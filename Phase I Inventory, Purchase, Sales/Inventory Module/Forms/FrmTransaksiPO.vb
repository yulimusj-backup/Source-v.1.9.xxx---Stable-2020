Option Explicit On
Option Strict Off

Public Class FrmTransaksiPO

    ''@rosadinnn
    ''@MUaT
    Dim tblCboUnit As New DseTransaksiPO.CboUnitDataTable
    Dim rowCboUnit As DseTransaksiPO.CboUnitRow
    Dim tblMultiUnit As New DseTransaksiPO.tblmaster_product_item_multiunitDataTable
    Dim rowMultiUnit As DseTransaksiPO.tblmaster_product_item_multiunitRow
    Dim objUnit As New Core_Module.ClsMasterUnit
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser

    Public UserCode As String = "USER"
    Public BAckColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public objCurrency As New Core_Module.clsMasterCurrency
    Public ObjCD As New Core_Module.ClsCekData

    Dim DefaultPrecision As Integer = 0
    Dim Buffer As String = ""
    Dim baru As Boolean = True
    Dim IPONo As String = ""
    Dim ObjPO As New ClsTransaksiPO
    Dim ObjSupplier As New Core_Module.clsMasterSupplier
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjUser As New Core_Module.ClsUser

    'terkait quick po
    Dim LangsungClose As Boolean = False

    ''@rosadinnn
    ''@MUaT
    Dim TBL As New DseTransaksiPO.DataTable9DataTable
    Public Enum EnumTolerance As Byte
        TOLERANCE_DOWN
        TOLERANCE_UP
    End Enum

    Dim ObjItem As New ClsTransaksiPOItem
    Dim ObjLicense As New Core_Module.ClsLicense
    Dim ObjCompanyData As New Core_Module.ClsMasterCompanyData
    Dim TBL_Add As New Core_Module.DseMasterCompanyData.tbltrans_company_data_itemDataTable
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
    Dim vrowItemDelivery As DseTransaksiPO.tbltrans_po_itemdeliveryRow
    Public vTBLItemDelivery As New DseTransaksiPO.tbltrans_po_itemdeliveryDataTable

    'Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Me.Close()
    'End Sub

    Dim ObjSetting As New Core_Module.ClsTransaksiSetting

    Private Sub chkInsertFromIPO_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkInsertFromIPO.CheckedChanged
        If Me.chkInsertFromIPO.Checked = True Then
            Me.SplitContainer1.Panel1Collapsed = False
            Me.ObjCD.KendalikanControl(Me.gboItem.Controls, True, False, True, False, False, False, False, False)
            Me.cboFilterIPO.SelectedIndex = 0
            Me.btnKanan.Enabled = True
            Me.btnKananSemua.Enabled = True
            Me.dgvItemIPO.ReadOnly = True
        Else
            Me.SplitContainer1.Panel1Collapsed = True
            Me.ObjCD.KendalikanControl(Me.gboItem.Controls, True, True, True, True, False, False, False, False)
            Me.cboFilterIPO.SelectedIndex = -1
            Me.btnKanan.Enabled = False
            Me.btnKananSemua.Enabled = False
            Me.TxtIPONo.Clear()
        End If
    End Sub

    Private Sub FrmTransaksiPO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim hasil As MsgBoxResult

        Try
            If Me.TxtRemark.Enabled = True Then
                Me.BringToFront()

                If Me.LangsungClose = False Then
                    hasil = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                Else
                    hasil = MsgBoxResult.No
                End If

                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPO.SimpanBaru(Me.dgvItem, Me.UserCode, Me, Me.vTBLItemDelivery) = True Then
                                Me.DialogResult = Windows.Forms.DialogResult.OK
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPO.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me, Me.vTBLItemDelivery) = True Then
                                Me.DialogResult = Windows.Forms.DialogResult.OK
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
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
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPO_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'Me.baru = True
            Me.SplitContainer1.Panel1Collapsed = True
            'Me.NonAktifkanDetail()

            ''@rosadinnn
            ''@MUaT
            UpdateKolomPO()

            Me.AktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.dgvItemIPO.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
            Me.ObjCompanyData.IsiCboAdd(Me.TBL_Add, Me.cboShipTo)
            Me.SetIndexCboShipToDefault()
            Control.CheckForIllegalCrossThreadCalls = False
            Me.dgvItem.Columns(Me.ColIPONo.Index).Visible = False
            Me.SplitContainer1.Visible = False
            Me.KosongkanItem()
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetIndexCboShipToDefault()
        Dim DefaultIdentifier As String = ""
        Dim row As Core_Module.DseMasterCompanyData.tbltrans_company_data_itemRow
        Dim DV As Data.DataRowView
        Dim i As Integer = 0

        Try
            DefaultIdentifier = Me.ObjCompanyData.GetIdentifierShipToDefault

            For Each DV In Me.TBL_Add.DefaultView
                row = DV.Row
                If row.Add_Identifier = DefaultIdentifier Then
                    Me.cboShipTo.SelectedIndex = i
                    Me.TxtIdentifierShipTo.Text = DefaultIdentifier
                    Exit Sub
                End If
                i += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.Enter
        Try
            Me.Buffer = Me.txtPrimarySupplier.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.GotFocus
        txtSupplierCode.BackColor = Me.BAckColorAll
    End Sub

    Private Sub txtSupplierCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.LostFocus
        txtSupplierCode.BackColor = Me.TxtPrimaryCodeNo.BackColor
    End Sub

    Private Sub txtContactPerson_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtContactPerson.GotFocus
        txtContactPerson.BackColor = Me.BAckColorAll
    End Sub

    Private Sub txtContactPerson_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtContactPerson.LostFocus
        txtContactPerson.BackColor = Me.TxtPrimaryCodeNo.BackColor
    End Sub

    Private Sub txtPhoneNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhoneNo.GotFocus
        txtPhoneNo.BackColor = Me.BAckColorAll
    End Sub

    Private Sub txtPhoneNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhoneNo.LostFocus
        txtPhoneNo.BackColor = Me.TxtPrimaryCodeNo.BackColor
    End Sub

    Private Sub TxtShipVia_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtShipVia.GotFocus
        TxtShipVia.BackColor = Me.BAckColorAll
    End Sub

    Private Sub TxtShipVia_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtShipVia.LostFocus
        TxtShipVia.BackColor = Me.TxtPrimaryCodeNo.BackColor
    End Sub

    Private Sub txtCurrencyCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.Enter
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.Buffer = Me.txtPrimaryCurrency.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.GotFocus
        txtCurrencyCode.BackColor = Me.BAckColorAll
    End Sub

    Private Sub txtCurrencyCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.LostFocus
        txtCurrencyCode.BackColor = Me.TxtPrimaryCodeNo.BackColor
    End Sub

    Private Sub txtPaymentTerm_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPaymentTerm.GotFocus
        txtPaymentTerm.BackColor = Me.BAckColorAll
    End Sub

    Private Sub txtPaymentTerm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPaymentTerm.LostFocus
        txtPaymentTerm.BackColor = Me.TxtPrimaryCodeNo.BackColor
    End Sub

    Private Sub txtPaymentDueDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPaymentDueDate.GotFocus
        txtPaymentDueDate.BackColor = Me.BAckColorAll
    End Sub

    Private Sub txtPaymentDueDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPaymentDueDate.LostFocus
        txtPaymentDueDate.BackColor = Me.TxtPrimaryCodeNo.BackColor
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        TxtRemark.BackColor = Me.BAckColorAll
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        TxtRemark.BackColor = Me.TxtPrimaryCodeNo.BackColor
    End Sub

    Private Sub TxtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.GotFocus
        TxtCodeNo.BackColor = Me.BAckColorAll
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.LostFocus
        TxtCodeNo.BackColor = Me.TxtPrimaryCodeNo.BackColor
    End Sub

    Private Sub TxtQtyperBatch_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.Enter
        Me.ObjCD.Number_Enter(Me.TxtQtyperBatch, sender, e)
    End Sub

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.GotFocus
        TxtQtyperBatch.BackColor = Me.BAckColorAll
    End Sub

    Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.LostFocus
        TxtQtyperBatch.BackColor = Me.TxtPrimaryCodeNo.BackColor
    End Sub

    Private Sub TxtUnitPrice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Enter
        Try
            Me.Buffer = Me.TxtUnitPrice.Text
            Me.ObjCD.Number_Enter(Me.TxtUnitPrice, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.GotFocus
        TxtUnitPrice.BackColor = Me.BAckColorAll
    End Sub

    Private Sub TxtUnitPrice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.LostFocus
        TxtUnitPrice.BackColor = Me.TxtPrimaryCodeNo.BackColor
    End Sub

    Private Sub TxtDiscount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Enter
        Try
            Me.Buffer = Me.TxtDiscount.Text
            Me.ObjCD.Number_Enter(Me.TxtDiscount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.GotFocus
        TxtDiscount.BackColor = Me.BAckColorAll
    End Sub

    Private Sub TxtDiscount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.LostFocus
        TxtDiscount.BackColor = Me.TxtPrimaryCodeNo.BackColor
    End Sub

    Private Sub TxtRemarkItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.GotFocus
        TxtRemarkItem.BackColor = Me.BAckColorAll
    End Sub

    Private Sub TxtRemarkItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.LostFocus
        TxtRemarkItem.BackColor = Me.TxtPrimaryCodeNo.BackColor
    End Sub

    Private Sub TxtIPONo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtIPONo.GotFocus
        TxtIPONo.BackColor = Me.BAckColorAll
    End Sub

    Private Sub TxtIPONo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtIPONo.LostFocus
        TxtIPONo.BackColor = Me.TxtPrimaryCodeNo.BackColor
    End Sub

    'Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim ij As New Inventory_Module.FrmCetakTransaksiPO

    '    Try
    '        If Me.TxtPrimaryPurchaseOrder.Text <> "" Then
    '            If IsNumeric(Me.TxtPrimaryPurchaseOrder.Text) = True Then
    '                If CDec(Me.TxtPrimaryPurchaseOrder.Text) >= 0 Then
    '                    'hanya PO yang sudah diotorisasi yang boleh diprint
    '                    If Me.ObjPO.GetAuthorizationLevel(Me.TxtPrimaryPurchaseOrder.Text) <> Me.ObjPO.GetAuthorizationLevelNeeded(Me.TxtPrimaryPurchaseOrder.Text) Then
    '                        MsgBox("This purchase order require more authorization, please check authorization needed.", _
    '                        MsgBoxStyle.Information, "Information")
    '                        Exit Sub
    '                    Else
    '                        If ObjPO.CetakPO(ij.crv1, Me.TxtPrimaryPurchaseOrder.Text) = True Then
    '                            ij.ShowDialog()
    '                        End If
    '                        Exit Sub
    '                    End If
    '                Else
    '                    Me.btnFind.PerformClick()
    '                End If
    '            Else
    '                Me.btnFind.PerformClick()
    '            End If
    '        Else
    '            Me.btnFind.PerformClick()
    '        End If

    '        If Me.TxtPrimaryPurchaseOrder.Text <> "" And Me.TxtPrimaryPurchaseOrder.Text <> "-1" Then
    '            'hanya PO yang sudah diotorisasi yang boleh diprint
    '            If Me.ObjPO.GetAuthorizationLevel(Me.TxtPrimaryPurchaseOrder.Text) <> Me.ObjPO.GetAuthorizationLevelNeeded(Me.TxtPrimaryPurchaseOrder.Text) Then
    '                MsgBox("This purchase order require more authorization, please check authorization needed.", _
    '                MsgBoxStyle.Information, "Information")
    '                Exit Sub
    '            Else
    '                If ObjPO.CetakPO(ij.crv1, Me.TxtPrimaryPurchaseOrder.Text) = True Then
    '                    ij.ShowDialog()
    '                End If
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ij.Dispose()
    '    End Try
    'End Sub

    Private Sub btnFindSupplier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSupplier.Click
        Dim i As New Core_Module.FrmCariMasterSupplier
        i.BackColorAll = Me.BAckColorAll
        Dim ApakahLokal As Boolean = False
        Dim Dummy As String = ""
        Dim Dummy2 As String = ""
        Dim IndexStr, IndexStr1 As Integer

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtSupplierCode.Text <> "" Then
                i.txtCari.Text = Me.txtSupplierCode.Text
            End If

            i.ShowDialog()
            If i.PrimarySupplier > 0 Then
                Me.txtPrimarySupplier.Text = i.PrimarySupplier
                Dim SupplierName, CPName, PhoneNo As String
                SupplierName = ""
                CPName = ""
                PhoneNo = ""

                Me.ObjSupplier.Get_SuppName_ContactPerson_PhoneNo(Me.txtPrimarySupplier.Text, SupplierName, CPName, PhoneNo)
                Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                Me.TxtSupplierName.Text = SupplierName
                Me.ObjSupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                Me.cboSupplierAddress.SelectedIndex = 0
                Me.cboSupplierAddress.Focus()
                Me.txtContactPerson.Text = CPName
                Me.txtPhoneNo.Text = PhoneNo
                Me.txtPrimaryCurrency.Text = Me.ObjSupplier.GetPrimaryDefaultCurrency(Me.txtPrimarySupplier.Text)
                Me.txtCurrencyCode.Text = Me.objCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                Me.SetDefaultPrecision()
                'ganti no transaksi
                If Me.baru = True Then
                    If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                        Me.TxtPONo.Text = Me.ObjPO.GetNoBaru1(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                    End If
                Else
                    ApakahLokal = Me.ObjSupplier.ApakahLokal(Me.txtPrimarySupplier.Text)
                    If ApakahLokal = True Then
                        IndexStr = Me.TxtPONo.Text.IndexOf("L")
                        IndexStr1 = Me.TxtPONo.Text.IndexOf("I", 12)
                    Else
                        IndexStr = Me.TxtPONo.Text.IndexOf("I", 12)
                        IndexStr1 = Me.TxtPONo.Text.IndexOf("L")
                    End If
                    If IndexStr = -1 Then
                        If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                            Me.TxtPONo.Text = Me.ObjPO.GetNoBaru1(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                        End If
                    End If
                End If

                Me.cboPaymentDueDateReference.SelectedIndex = 1
                Me.txtPaymentTerm.Text = Me.ObjSupplier.GetPaymentTerm(Me.txtPrimarySupplier.Text)
                Me.txtPaymentDueDate.Text = Me.txtPaymentTerm.Text
            Else
                Me.txtPrimarySupplier.Text = "-1"
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.cboSupplierAddress.Items.Clear()
                Me.txtContactPerson.Clear()
                Me.txtPhoneNo.Clear()
                Me.txtSupplierCode.Focus()
                Me.cboPaymentDueDateReference.SelectedIndex = 1
                Me.txtPaymentTerm.Text = 0
                Me.txtPaymentDueDate.Text = Me.txtPaymentTerm.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtSupplierCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSupplierCode.KeyDown
        Dim ApakahLokal As Boolean = False
        Dim Dummy As String = ""
        Dim Dummy2 As String = ""
        Dim IndexStr, IndexStr1 As Integer

        Try
            If e.KeyValue = 13 Then
                Dim Primary As Double = ObjSupplier.GetPrimary(Me.txtSupplierCode.Text)
                If Primary <> -1 Then
                    Me.txtPrimarySupplier.Text = Primary
                    Dim SupplierName, CPName, PhoneNo As String
                    SupplierName = ""
                    CPName = ""
                    PhoneNo = ""

                    Me.ObjSupplier.Get_SuppName_ContactPerson_PhoneNo(Me.txtPrimarySupplier.Text, SupplierName, CPName, PhoneNo)
                    Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                    Me.TxtSupplierName.Text = SupplierName
                    Me.ObjSupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                    Me.cboSupplierAddress.SelectedIndex = 0
                    Me.cboSupplierAddress.Focus()
                    Me.txtContactPerson.Text = CPName
                    Me.txtPhoneNo.Text = PhoneNo
                    Me.txtPrimaryCurrency.Text = Me.ObjSupplier.GetPrimaryDefaultCurrency(Me.txtPrimarySupplier.Text)
                    Me.txtCurrencyCode.Text = Me.objCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                    Me.SetDefaultPrecision()
                    'ganti no transaksi
                    If Me.baru = True Then
                        If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                            Me.TxtPONo.Text = Me.ObjPO.GetNoBaru1(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                        End If
                    Else
                        ApakahLokal = Me.ObjSupplier.ApakahLokal(Me.txtPrimarySupplier.Text)
                        If ApakahLokal = True Then
                            IndexStr = Me.TxtPONo.Text.IndexOf("L")
                            IndexStr1 = Me.TxtPONo.Text.IndexOf("I", 12)
                        Else
                            IndexStr = Me.TxtPONo.Text.IndexOf("I", 12)
                            IndexStr1 = Me.TxtPONo.Text.IndexOf("L")
                        End If
                        If IndexStr = -1 Then
                            If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                                Me.TxtPONo.Text = Me.ObjPO.GetNoBaru1(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                            End If
                        End If
                    End If
                    Me.cboPaymentDueDateReference.SelectedIndex = 1
                    Me.txtPaymentTerm.Text = Me.ObjSupplier.GetPaymentTerm(Me.txtPrimarySupplier.Text)
                    Me.txtPaymentDueDate.Text = Me.txtPaymentTerm.Text
                Else
                    Me.btnFindSupplier.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSupplierCode.Leave
        Dim ApakahLokal As Boolean = False
        Dim Dummy As String = ""
        Dim Dummy2 As String = ""
        Dim SupplierName, CPName, PhoneNo As String
        SupplierName = ""
        CPName = ""
        PhoneNo = ""
        Dim IndexStr, IndexStr1 As Integer

        Try
            If Me.txtSupplierCode.Text <> "" Then
                Dim Primary As Double = ObjSupplier.GetPrimary(Me.txtSupplierCode.Text)
                If Primary <> -1 Then
                    Me.txtPrimarySupplier.Text = Primary
                    Me.ObjSupplier.Get_SuppName_ContactPerson_PhoneNo(Me.txtPrimarySupplier.Text, SupplierName, CPName, PhoneNo)
                    Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                    Me.TxtSupplierName.Text = SupplierName
                    Me.ObjSupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                    Me.cboSupplierAddress.SelectedIndex = 0
                    Me.txtContactPerson.Text = CPName
                    Me.txtPhoneNo.Text = PhoneNo
                    Me.txtPrimaryCurrency.Text = Me.ObjSupplier.GetPrimaryDefaultCurrency(Me.txtPrimarySupplier.Text)
                    Me.txtCurrencyCode.Text = Me.objCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                    Me.SetDefaultPrecision()
                    'ganti no transaksi
                    If Me.baru = True Then
                        If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                            Me.TxtPONo.Text = Me.ObjPO.GetNoBaru1(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                        End If
                    Else
                        ApakahLokal = Me.ObjSupplier.ApakahLokal(Me.txtPrimarySupplier.Text)
                        If ApakahLokal = True Then
                            IndexStr = Me.TxtPONo.Text.IndexOf("L")
                            IndexStr1 = Me.TxtPONo.Text.IndexOf("I", 12)
                        Else
                            IndexStr = Me.TxtPONo.Text.IndexOf("I", 12)
                            IndexStr1 = Me.TxtPONo.Text.IndexOf("L")
                        End If
                        If IndexStr = -1 Then
                            If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                                Me.TxtPONo.Text = Me.ObjPO.GetNoBaru1(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                            End If
                        End If
                    End If

                    If Me.Buffer <> Me.txtPrimarySupplier.Text Then
                        Me.cboPaymentDueDateReference.SelectedIndex = 1
                        Me.txtPaymentTerm.Text = Me.ObjSupplier.GetPaymentTerm(Me.txtPrimarySupplier.Text)
                        Me.txtPaymentDueDate.Text = Me.txtPaymentTerm.Text
                    End If
                Else
                    Me.btnFindSupplier.PerformClick()
                End If
            Else
                Me.txtPrimarySupplier.Text = "-1"
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.cboSupplierAddress.Items.Clear()
                Me.txtContactPerson.Clear()
                Me.txtPhoneNo.Clear()
                Me.cboPaymentDueDateReference.SelectedIndex = 1
                Me.txtPaymentTerm.Text = 0
                Me.txtPaymentDueDate.Text = Me.txtPaymentTerm.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboSupplierAddress_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSupplierAddress.SelectedIndexChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.cboSupplierAddress.SelectedIndex <> -1 Then
                    Me.txtSupplierAddIdentifier.Text = Me.ObjSupplier.GetIdentifierAdd(Me.txtPrimarySupplier.Text, _
                    Me.cboSupplierAddress.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAddress_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddress.LinkClicked
        Dim ij As New Core_Module.FrmDetailAddressSupplier

        Try
            If Me.txtPrimarySupplier.Text = "" Or Me.txtPrimarySupplier.Text = "-1" Then
                If Me.TxtRemark.Enabled = True Then
                    MsgBox("Please specify supplier first.", MsgBoxStyle.Information, "Information")
                    Me.txtSupplierCode.Focus()
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If

            ij.vAddIdentifier = Me.txtSupplierAddIdentifier.Text
            ij.PrimarySup = Me.txtPrimarySupplier.Text
            If Me.TxtRemark.Enabled = True Then
                ij.btnOK.BringToFront()
            Else
                ij.cboCustAddCode.Enabled = False
            End If
            ij.AcceptButton = ij.btnOK
            ij.btnExit.TabStop = False
            ij.ShowDialog()
            If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                Me.ObjSupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                Me.cboSupplierAddress.Text = Me.ObjSupplier.GetKodeAddress(ij.vAddIdentifierResult)
            End If
            ij.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAuthorization_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAuthorization.LinkClicked
        If Me.TxtPrimaryPurchaseOrder.Text = "" Or Me.TxtPrimaryPurchaseOrder.Text = "-1" _
        And Me.TxtRemark.Enabled = False Then
            Exit Sub
        End If

        Dim i As New FrmTransaksiSupplierAuthorizationDetail
        Dim jmlbaris As Integer = 0
        Dim row As DataGridViewRow
        Try
            i.dgvDetail.Rows.Clear()
            'tambahkan 3 baris authorization
            jmlbaris = Me.ObjPO.AuthorizationNeeded
            'masukkan baris dan detail data otorisasi
            For j As Integer = 1 To jmlbaris
                i.dgvDetail.Rows.Add()
                row = i.dgvDetail.Rows(i.dgvDetail.Rows.Count - 1)
                If row.Index = 0 Then
                    row.Cells(1).Value = "Purchase Staff Authorization"
                    If Me.ObjPO.PrimaryPurchaseStaffAuthorization <> -1 Then
                        row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjPO.PrimaryPurchaseStaffAuthorization)
                        row.Cells(3).Value = Format(Me.ObjPO.PurchaseStaffAuthorizationDate, "dd MMM yyyy - HH:mm")
                    Else
                        row.Cells(2).Value = "-"
                        row.Cells(3).Value = "-"
                    End If
                ElseIf row.Index = 1 Then
                    row.Cells(1).Value = "Purchase Manager Authorization"
                    If Me.ObjPO.PrimaryManagerStaffAuthorization <> -1 Then
                        row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjPO.PrimaryManagerStaffAuthorization)
                        row.Cells(3).Value = Format(Me.ObjPO.PurchaseManagerAuthorizationDate, "dd MMM yyyy - HH:mm")
                    Else
                        row.Cells(2).Value = "-"
                        row.Cells(3).Value = "-"
                    End If
                ElseIf row.Index = 2 Then
                    row.Cells(1).Value = "Director Authorization"
                    If Me.ObjPO.PrimaryDirectorStaffAuthorization <> -1 Then
                        row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjPO.PrimaryDirectorStaffAuthorization)
                        row.Cells(3).Value = Format(Me.ObjPO.PurchaseDirectorAuthorizationDate, "dd MMM yyyy - HH:mm")
                    Else
                        row.Cells(2).Value = "-"
                        row.Cells(3).Value = "-"
                    End If
                End If
            Next
            'tunjukkan form
            Me.ObjCD.SetNumber(i.dgvDetail)
            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i = Nothing
        End Try
    End Sub

    Private Function GetJmlItemDelivery(ByVal vIdentifier As String) As Double
        Dim Retval As Integer = 0

        Try
            For Each Me.vrowItemDelivery In Me.vTBLItemDelivery.Rows
                If Me.vrowItemDelivery.identifier = vIdentifier Then
                    Retval += 1
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub Hitung_Total()
        Dim row As DataGridViewRow
        Dim SubTotal As Decimal = 0

        Try
            'hitung sub total
            For Each row In Me.dgvItem.Rows
                SubTotal += CDec(row.Cells(Me.ColAmount.Index).Value)
            Next
            Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision))
            'hitung pajak
            If Me.chkTax.Checked = True Then
                If Me.txtPrimaryCurrency.Text = 1 Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Me.txtTax.Text * SubTotal / 100, _
                    0, Core_Module.ClsCekData.EnumTipeRounding.Down), DefaultPrecision)
                Else
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTax.Text * SubTotal / 100, _
                    DefaultPrecision, Core_Module.ClsCekData.EnumTipeRounding.Down)
                End If
            Else
                Me.txtTax.Text = "0"
                Me.txtTaxAmount.Text = "0"
            End If
            'hitung total = sub total + tax
            Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.txtSubTotal.Text) + _
            CDec(Me.txtTaxAmount.Text), DefaultPrecision))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetDefaultPrecision()
        Dim row As DataGridViewRow
        Dim DummyDefPrec As Integer = 0

        Try
            DummyDefPrec = Me.DefaultPrecision
            Me.DefaultPrecision = Me.objCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)

            If DummyDefPrec <> Me.DefaultPrecision Then
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

                'yang diformat dgv-nya:
                For Each row In Me.dgvItem.Rows
                    'digit grouping: qty, unit price, nett price, amount
                    row.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColUnitPrice.Index).Value, Me.DefaultPrecision))
                    row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColUnitPrice.Index).Value) * (100 - CDec(row.Cells(Me.ColDiscount.Index).Value)) / 100, _
                    Me.DefaultPrecision)
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(row.Cells(Me.ColQtyMulti.Index).Value), DefaultPrecision)
                Next

                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

                'hitung ulang yang sudah ada didalam DGV
                Me.Hitung_Total()
                'hitung ulang yang ada didalam gbo input
                Me.Hitung_Amount()
                Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotal.Text, DefaultPrecision))
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtTaxAmount.Text, DefaultPrecision))
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.TxtTotal.Text, DefaultPrecision))
                Me.txtAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtAmount.Text, DefaultPrecision))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCurrency.Click
        Dim i As New Core_Module.FrmCariMasterCurrency
        i.BackColorAll = Me.BAckColorAll
        Try
            If Me.objCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                Me.chkTax.Focus()
                Me.txtPrimaryCurrency.Text = objCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                Me.SetDefaultPrecision()
            Else
                i.cboCriteria.SelectedIndex = 0
                If Me.txtCurrencyCode.Text <> "" Then
                    i.txtCari.Text = Me.txtCurrencyCode.Text
                End If

                i.ShowDialog()
                If i.PrimaryCurrency > 0 Then
                    Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                    Me.txtCurrencyCode.Text = objCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                    Me.SetDefaultPrecision()
                    Me.chkTax.Focus()
                Else
                    Me.txtCurrencyCode.Clear()
                    Me.txtCurrencyCode.Focus()
                    Me.SetDefaultPrecision()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub chkTax_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTax.CheckedChanged
        Dim Dummy As String = ""
        Try
            If Me.chkTax.Checked = True Then
                Me.txtTax.ReadOnly = False
                Me.txtTax.TabStop = True
            Else
                Me.txtTax.ReadOnly = True
                Me.txtTax.TabStop = False
                Me.txtTax.Text = 0
            End If

            If Me.TxtRemark.Enabled = True Then
                If Me.chkTax.Checked = True Then
                    Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                    If Dummy = "" Then
                        Dummy = "0"
                    End If
                    Me.txtTax.Text = Me.ObjCD.DigitGrouping(Dummy, 2)

                    'If CDbl(Me.txtTax.Text) = 0 Then
                    '    Me.txtTax.Text = Me.ObjCD.DigitGrouping("10", 2)
                    'End If
                End If

                Me.HitungPajak()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPaymentDueDate_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPaymentDueDate.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtPaymentDueDate, sender, e, False)
    End Sub

    Private Sub txtPaymentDueDate_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPaymentDueDate.KeyDown
        Try
            'harus angka bulat
            If e.KeyCode = Keys.OemPeriod Or e.KeyCode = Keys.Decimal Then
                e.SuppressKeyPress = True
                Exit Sub
            End If
            Me.ObjCD.Number_KeyDown(Me.txtPaymentDueDate, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNo.Click
        Dim i As New Core_Module.frmCariMasterProduct
        Try
            'cek apakah sudah isi kode currency
            i.BackColorAll = Me.BAckColorAll
            If Me.txtPrimarySupplier.Text <> "" Then
                If Me.txtPrimarySupplier.Text = -1 And Me.txtPrimarySupplier.Text <> "" Then
                    '             Me.TabControl1.SelectedIndex = 1
                    Me.txtSupplierCode.Focus()
                    MsgBox("To start entering item data, you specify supplier data first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.txtPrimaryCurrency.Text <> "" Then
                If Me.txtPrimaryCurrency.Text = -1 And Me.txtPrimaryCurrency.Text <> "" Then
                    '            Me.TabControl1.SelectedIndex = 0
                    Me.txtCurrencyCode.Focus()
                    MsgBox("To start entering item data, you specify currency data first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                Me.dtpDeliveryDate.Focus()
                Me.SetGbo()

                ''@rosadinnn
                ''@MUaT
                Me.PrepareMultiUnit(Me.TxtPrimaryCodeNo.Text)

            Else
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtPrimaryCategory.Text = "-1"
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub KosongkanItem()
        Dim CheckDummy As Boolean = False
        Dim DummyDate As Date = Today

        Try
            CheckDummy = Me.dtpDeliveryDate.Checked
            DummyDate = Me.dtpDeliveryDate.Value

            ObjCD.KosongkanControl(Me.gboItem.Controls, True, True, False)

            ''@rosadinnn
            ''@MUaT
            ObjCD.KosongkanControl(Me.SplitContainer3.Panel1.Controls, True, True, False)
            ObjCD.KosongkanControl(Me.SplitContainer3.Panel2.Controls, True, True, False)

            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryUnits.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.txtPrimaryIPO.Text = "-1"
            Me.TxtQtyperBatch.Text = "0"
            Me.TxtUnitPrice.Text = "0"
            Me.TxtDiscount.Text = "0"
            Me.TxtDiscAmount.Text = 0
            Me.TxtNettPrice.Text = 0
            Me.TxtDiscFromPerct.Text = "1"
            Me.txtAmount.Text = "0"
            Me.lblUnits.Text = "Units"
            Me.LblUnitPrice.Text = "Unit Price"
            Me.llbHistoricalPrice.Enabled = False

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiPO_Override_Promise_To_Deliver") = False Then
                Me.dtpDeliveryDate.Checked = False
                Me.dtpDeliveryDate.Value = Me.dtpDate.Value
            Else
                Me.dtpDeliveryDate.Checked = CheckDummy
                If Me.dtpDeliveryDate.Checked = True Then
                    Me.dtpDeliveryDate.Value = DummyDate
                End If
            End If

            ''@rosadinnn
            ''@MUaT
            Me.lblUnits.Visible = True
            Me.CboMultiUnit.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetGbo()
        Dim UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc As Decimal
        Try
            Me.lblUnits.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
            Me.txtPrimaryUnits.Text = Me.ObjProduct.GetPrimaryUnit(Me.TxtPrimaryCodeNo.Text)
            Me.chkUseSecondary.Checked = Me.ObjProduct.ApakahUseSecondary(Me.TxtPrimaryCodeNo.Text)
            If Me.chkUseSecondary.Checked = True Then
                Me.LblUnitPrice.Text = "Unit Price/" + Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text)
            Else
                Me.LblUnitPrice.Text = "Unit Price/" + Me.lblUnits.Text
            End If
            'cari historical price
            If Me.TxtPrimaryCodeNo.Text <> "" Or Me.TxtPrimaryCodeNo.Text <> "-1" Then
                Me.llbHistoricalPrice.Enabled = True

                ''@rosadinnn
                ''@MUaT
                ''@tambah parameter unit
                Me.ObjPO.GetHistoricalPrice(Me.txtPrimaryUnits.Text, Me.TxtPrimaryCodeNo.Text, _
                Me.txtPrimaryCurrency.Text, UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc)

                Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.DefaultPrecision)
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                Me.TxtNettPrice.Text = Me.ObjCD.DigitGrouping(NettPrice, Me.DefaultPrecision)
                Me.TxtDiscAmount.Text = Me.ObjCD.DigitGrouping(DiscAmount, Me.DefaultPrecision)
                Me.TxtDiscFromPerct.Text = DiscFromPerc
            Else
                Me.llbHistoricalPrice.Enabled = False
            End If
            Me.Hitung_Amount()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Hitung_Amount()
        Try
            If Me.TxtDiscFromPerct.Text = "" Then
                Me.TxtDiscFromPerct.Text = "1"
            End If
            If Me.TxtDiscFromPerct.Text = "1" Then
                'hitung disc amount
                Me.TxtDiscAmount.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtUnitPrice.Text) * CDec(Me.TxtDiscount.Text)) / 100, DefaultPrecision)
            Else
                'hitung disc %
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtDiscAmount.Text) * 100) / CDec(Me.TxtUnitPrice.Text), 2)
            End If
            'hitung nett price
            Me.TxtNettPrice.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtUnitPrice.Text) - CDec(Me.TxtDiscAmount.Text), _
            DefaultPrecision)

            'hitung amount
            If Me.chkUseSecondary.Checked = False Then
                Me.txtAmount.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtNettPrice.Text) * _
                CDec(Me.TxtQtyperBatch.Text), DefaultPrecision)
            Else
                Me.txtAmount.Text = Me.ObjCD.DigitGrouping(0, DefaultPrecision)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatch.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatch, sender, e)
    End Sub

    Private Sub TxtQtyperBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatch.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatch, sender, e)
    End Sub

    Private Sub TxtUnitPrice_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPrice.KeyDown
        Me.Buffer = Me.TxtUnitPrice.Text
        Me.ObjCD.Number_KeyDown(Me.TxtUnitPrice, sender, e)
    End Sub

    Private Sub TxtUnitPrice_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPrice.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtUnitPrice, sender, e)
    End Sub

    Private Sub TxtDiscount_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDiscount, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtDiscount, sender, e)
    End Sub

    Private Sub txtTax_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.Enter
        Try
            Me.ObjCD.Number_Enter(Me.txtTax, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.GotFocus
        Try
            If Me.txtTax.ReadOnly = False And Me.txtTax.Enabled = True Then
                Me.txtTax.BackColor = Me.BAckColorAll
            Else
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTax.KeyDown
        Try
            If Me.txtTax.ReadOnly = False Then
                Me.ObjCD.Number_KeyDown(Me.txtTax, sender, e, 100)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTax.KeyUp
        Try
            If Me.txtTax.ReadOnly = False Then
                Me.ObjCD.Number_KeyUp(Me.txtTax, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.Leave
        Try
            Me.txtTax.Text = Me.ObjCD.DigitGrouping(Me.txtTax.Text, 2)
            If Me.TxtRemark.Enabled = True And IsNumeric(Me.txtTax.Text) = True And _
            IsNumeric(Me.txtSubTotal.Text) Then
                Me.HitungPajak()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.LostFocus
        Try
            If Me.txtTax.Enabled = True And Me.txtTax.ReadOnly = False Then
                Me.txtTax.BackColor = Me.TxtPrimaryPurchaseOrder.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungPajak()
        Try
            If CDec(Me.txtTax.Text) <= 100 Then
                'hitung pajak
                If Me.chkTax.Checked = True Then
                    If Me.txtPrimaryCurrency.Text = "1" Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Me.txtTax.Text * CDec(Me.txtSubTotal.Text) / 100, _
                        0, Core_Module.ClsCekData.EnumTipeRounding.Down), DefaultPrecision)
                    Else
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTax.Text * CDec(Me.txtSubTotal.Text) / 100, _
                        DefaultPrecision, Core_Module.ClsCekData.EnumTipeRounding.Down)
                    End If
                Else
                    Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.objCurrency.GetDecimal_DefaultCurrency)
                End If
                'hitung total = sub total + tax
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.txtSubTotal.Text) _
                + CDec(Me.txtTaxAmount.Text), DefaultPrecision))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyCode.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Dim primary As Double = objCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                If primary <> -1 Then
                    Me.txtPrimaryCurrency.Text = primary
                    Me.txtCurrencyCode.Text = objCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                    Me.SetDefaultPrecision()
                    Me.chkTax.Focus()
                Else
                    Me.btnFindCurrency.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.Leave
        Dim i As New Core_Module.FrmCariMasterCurrency
        i.BackColorAll = Me.BAckColorAll

        Try
            If Me.txtCurrencyCode.Text <> "" Then
                Dim primary As Double = objCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                If primary <> -1 Then
                    Me.txtPrimaryCurrency.Text = primary
                    Me.txtCurrencyCode.Text = objCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                    Me.SetDefaultPrecision()
                Else
                    If Me.objCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                        Me.txtPrimaryCurrency.Text = objCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                        Me.SetDefaultPrecision()
                    Else
                        i.cboCriteria.SelectedIndex = 0
                        If Me.txtCurrencyCode.Text <> "" Then
                            i.txtCari.Text = Me.txtCurrencyCode.Text
                        End If

                        i.ShowDialog()
                        If i.PrimaryCurrency > 0 Then
                            Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                            Me.txtCurrencyCode.Text = objCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                            Me.SetDefaultPrecision()
                        Else
                            Me.txtCurrencyCode.Clear()
                            Me.txtCurrencyCode.Focus()
                            Me.SetDefaultPrecision()
                        End If
                    End If
                End If

                If Me.txtPrimaryCurrency.Text <> Me.Buffer Then
                    Me.HitungPajak()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i = Nothing
        End Try
    End Sub

    Public Function GetIndex(ByVal vPrimaryProduct As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    Return row.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndex(ByVal vPrimaryProduct As Double, ByVal vPrimaryIPO As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                row.Cells(Me.ColPrimaryIPO.Index).Value = vPrimaryIPO Then
                    Return row.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataItemBenar() As Boolean
        Try
            'data item salah bila:
            '1. product code yang diisikan salah
            If Me.txtDesc.Text = "" Then
                MsgBox("Please enter code no.", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNo.Text = ""
                Me.TxtCodeNo.Focus()
                Return False
            End If
            '2. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatch.Text) = False Then
                MsgBox("Please enter a positive number for quantity data.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatch.Focus()
                Return False
            End If
            '3. Bila sudah ada product pada dgv dan unit price, discount, nett price-nya, 
            'delivery datenya beda - saat ini satu product hanya bisa muncul satu kali pada 1 so
            Dim NettPrice As Decimal = 0
            NettPrice = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.TxtUnitPrice.Text) * (100 - CDec(Me.TxtDiscount.Text)) / 100, _
            DefaultPrecision))

            If Me.GetIndex(Me.TxtPrimaryCodeNo.Text) <> -1 Then
                If Me.GetIndex(Me.TxtPrimaryCodeNo.Text, Me.dtpDeliveryDate.Checked, Me.dtpDeliveryDate.Value, _
                Me.TxtUnitPrice.Text, NettPrice, Me.TxtDiscount.Text) = -1 Then
                    MsgBox(Me.TxtCodeNo.Text + " is already enterred, to enter 1 product code several time in one purchase order the unit price, discount, and delivery date must identical.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            '4. bila janji akan ada tanggal pengiriman dan kemudian tanggal pengirimannya dibelakang hari ini
            '   berikan peringatan
            If Me.dtpDeliveryDate.Checked = True And Me.dtpDeliveryDate.Value.Date < Today Then
                If MsgBox("Delivery date is smaller than transaction date, continue?", MsgBoxStyle.YesNo, _
                "Confirm") = MsgBoxResult.No Then
                    Return False
                End If
            End If
            '5. nett price tidak bisa 0
            If NettPrice = 0 Then
                MsgBox("Nett price cannot be 0", MsgBoxStyle.Information, "Information")
                Me.TxtUnitPrice.Focus()
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndex(ByVal vPrimaryProduct As Double, ByVal DeliveryDateCont As Boolean, ByVal DeliveryDate As Date, _
   ByVal UnitPrice As Decimal, ByVal NettPrice As Decimal, ByVal Discount As Decimal) As Integer
        Dim row As DataGridViewRow
        Try
            For Each row In Me.dgvItem.Rows
                If CBool(row.Cells(Me.ColDeliveryCont.Index).Value) = True Then

                    ''@rosadinnn
                    ''@MUaT
                    ''perubahan membandingkan delivery date
                    If IsDate(row.Cells(Me.ColDeliveryDate.Index).Value) = True Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And CBool(row.Cells(Me.ColDeliveryCont.Index).Value) = DeliveryDateCont And _
                                            CDate(row.Cells(Me.ColDeliveryDate.Index).Value) = DeliveryDate.Date And row.Cells(Me.ColUnitPrice.Index).Value = UnitPrice And _
                                            row.Cells(Me.ColDiscount.Index).Value = Discount And row.Cells(Me.ColNettPrice.Index).Value = NettPrice Then
                            Return row.Index
                        End If
                    End If
                    
                Else
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And CBool(row.Cells(Me.ColDeliveryCont.Index).Value) = DeliveryDateCont And _
                    row.Cells(Me.ColUnitPrice.Index).Value = UnitPrice And _
                    row.Cells(Me.ColDiscount.Index).Value = Discount And row.Cells(Me.ColNettPrice.Index).Value = NettPrice Then
                        Return row.Index
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    ''@update->perubahan col qty
    Public Sub Hitung_Amount(ByVal row As DataGridViewRow)
        Dim DiscPerct As Integer = 0

        Try
            'hitung nett price
            DiscPerct = row.Cells(Me.ColDiscFromPerct.Index).Value
            If DiscPerct = 1 Then
                row.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(row.Cells(Me.ColUnitPrice.Index).Value) * CDec(row.Cells(Me.ColDiscount.Index).Value))) / 100, DefaultPrecision))
            Else
                row.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(row.Cells(Me.ColDiscAmount.Index).Value) / CDec(row.Cells(Me.ColUnitPrice.Index).Value))) * 100, 2))
            End If
            row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
            CDec(row.Cells(Me.ColUnitPrice.Index).Value) - CDec(row.Cells(Me.ColDiscAmount.Index).Value) _
            , DefaultPrecision))

            'hitung amount
            If row.Cells(Me.colUserSecondaryUnits.Index).Value = False Then
                'row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColQty.Index).Value * _
                'row.Cells(Me.ColNettPrice.Index).Value, DefaultPrecision))
                ''@rosadinnn
                ''@MUaT
                row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColQtyMulti.Index).Value * _
                row.Cells(Me.ColNettPrice.Index).Value, DefaultPrecision))
            Else
                row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(0, DefaultPrecision))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInput.Click
        Dim Index1 As Integer = -1
        Dim rowDGV As DataGridViewRow
        Dim rowDelivery As DseTransaksiPO.tbltrans_po_itemdeliveryRow

        Try
            If Me.ApakahDataItemBenar = True Then
                'Dim NettPrice As Decimal = 0
                'NettPrice = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.TxtUnitPrice.Text) * _
                '(100 - CDec(Me.TxtDiscount.Text)) / 100, _
                'DefaultPrecision))

                Index1 = Me.GetIndex(Me.TxtPrimaryCodeNo.Text)
                If Index1 <> -1 Then
                    Index1 = Me.GetIndex(Me.TxtPrimaryCodeNo.Text, -1)
                End If

                'cari data yang sama persis
                If Index1 <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index1)

                    ''@rosadinnn
                    ''@MUaT
                    row.Cells(Me.ColQtyMulti.Index).Value = CDec(row.Cells(Me.ColQtyMulti.Index).Value) + CDec(Me.TxtQtyperBatch.Text)
                    Me.SetColQtyDefault(row)

                    Me.Hitung_Amount(row)
                    Me.Hitung_Total()

                    'tambahkan qty delivery pada baris terakhir
                    If row.Cells(Me.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                        Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + row.Cells(Me.ColIdentifier.Index).Value.ToString + "'"
                        Me.vTBLItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                        Me.vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(0).Row

                        ''@rosadinnn
                        ''@MUaT
                        Me.vrowItemDelivery.delivery_qty_multi += CDec(Me.TxtQtyperBatch.Text)
                        Me.vrowItemDelivery.delivery_qty = row.Cells(Me.ColQty.Index).Value

                    End If
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Me.dtpDeliveryDate.Checked = False Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.dtpDeliveryDate.Value.Date, "dd MMM yyyy")
                    End If
                    'cegah jangan sampe primary IPO kosong sama sekali
                    If Me.txtPrimaryIPO.Text = "" Then
                        Me.txtPrimaryIPO.Text = "-1"
                    End If
                    Me.dgvItem.Rows.Add()
                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColNo.Index).Value = Me.dgvItem.RowCount
                        .Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                        .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                        .Cells(Me.ColIPONo.Index).Value = Me.TxtIPONo.Text
                        .Cells(Me.ColQty.Index).Value = Me.TxtQtyperBatch.Text
                        .Cells(Me.ColUnit.Index).Value = Me.lblUnits.Text
                        .Cells(Me.ColUnitPrice.Index).Value = Me.TxtUnitPrice.Text
                        .Cells(Me.ColDiscount.Index).Value = Me.TxtDiscount.Text
                        .Cells(Me.ColNettPrice.Index).Value = Me.TxtNettPrice.Text
                        .Cells(Me.ColDiscAmount.Index).Value = Me.TxtDiscAmount.Text
                        .Cells(Me.ColDiscFromPerct.Index).Value = Me.TxtDiscFromPerct.Text
                        .Cells(Me.ColDeliveryCont.Index).Value = Me.dtpDeliveryDate.Checked
                        .Cells(Me.ColDeliveryDate.Index).Value = ExpiredDateStr
                        .Cells(Me.ColRemark.Index).Value = Me.TxtRemarkItem.Text
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text
                        .Cells(Me.ColPrimaryUnits.Index).Value = Me.txtPrimaryUnits.Text
                        .Cells(Me.ColPrimaryIPO.Index).Value = Me.txtPrimaryIPO.Text
                        .Cells(Me.colUserSecondaryUnits.Index).Value = Me.chkUseSecondary.Checked
                        .Cells(Me.ColIdentifier.Index).Value = Me.GetIdentifierBaru
                        .Cells(Me.ColAmount.Index).Value = Me.txtAmount.Text

                        ''@rosadinnn
                        ''@MUaT
                        .Cells(Me.ColPrimaryUnitMulti.Index).Value = Me.txtPrimaryUnits.Text
                        .Cells(Me.ColUnitCodeMulti.Index).Value = Me.lblUnits.Text
                        '---------------------------------------------------------
                        .Cells(Me.ColQtyMulti.Index).Value = Me.TxtQtyperBatch.Text
                        '---------------------------------------------------------
                        .Cells(Me.ColPrimaryUnits.Index).Value = Me.ObjProduct.GetPrimaryUnit(.Cells(Me.ColPrimaryCodeNo.Index).Value)
                        .Cells(Me.ColUnit.Index).Value = Me.objUnit.GetUnitCode(.Cells(Me.ColPrimaryUnits.Index).Value)
                        '---------------------------------------------------------
                        .Cells(Me.ColRatioDefault.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjPO.GetRatioDefault(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryUnitMulti.Index).Value), _
                                                                Me.ObjProduct.GetAngkaDecimal(.Cells(Me.ColPrimaryCodeNo.Index).Value))
                        .Cells(Me.ColRatioReference.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjPO.GetRatioReference(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryUnitMulti.Index).Value), _
                                                                ObjProduct.GetAngkaDecimal(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryUnitMulti.Index).Value))
                        '---------------------------------------------------------
                        Me.SetRatioAndQtyDefault(rowDGV, True)
                        .Cells(Me.ColToleranceDown.Index).Value = Me.ObjCD.DigitGrouping("0", ObjProduct.GetAngkaDecimal(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryUnitMulti.Index).Value))
                        .Cells(Me.ColToleranceUp.Index).Value = Me.ObjCD.DigitGrouping("0", ObjProduct.GetAngkaDecimal(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryUnitMulti.Index).Value))
                        .Cells(Me.ColToleranceDownPerc.Index).Value = Me.ObjCD.DigitGrouping("0", 2)
                        .Cells(Me.ColToleranceUpPerc.Index).Value = Me.ObjCD.DigitGrouping("0", 2)

                    End With

                    'tambahkan item delivery
                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    If CBool(rowDGV.Cells(Me.ColDeliveryCont.Index).Value) = True Then
                        rowDelivery = Me.vTBLItemDelivery.Newtbltrans_po_itemdeliveryRow
                        With rowDelivery
                            .delivery_date = Me.dtpDeliveryDate.Value
                            .delay = Me.ObjItem.GetPreparationDelay(Me.TxtPrimaryCodeNo.Text)
                            .delivery_qty = Me.TxtQtyperBatch.Text
                            .identifier = Me.dgvItem.Item(Me.ColIdentifier.Index, Me.dgvItem.RowCount - 1).Value
                            .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                            .primary_transaksi = Me.TxtPrimaryPurchaseOrder.Text
                            .remarks = ""

                            ''@rosadinnn
                            ''@MUaT
                            .delivery_qty_multi = rowDGV.Cells(Me.ColQtyMulti.Index).Value
                            .delivery_qty = rowDGV.Cells(Me.ColQty.Index).Value

                        End With
                        vTBLItemDelivery.Rows.Add(rowDelivery)
                    End If
                End If

                'kosongkan gbo item
                Me.KosongkanItem()
                Me.txtAmount.Text = Me.ObjCD.DigitGrouping(Me.txtAmount.Text, Me.DefaultPrecision)
                Me.TxtCodeNo.Focus()
                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                'hitung total
                Me.Hitung_Total()
            End If

            ''@rosadinnn
            ''@MUaT
            Me.FormatDgv()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetIdentifierBaru() As String
        Dim IdentifierD As Integer = 0
        Dim NilaiMax As Integer = 0
        Dim Retval As String = ""

        Try
            For Each rowDGV As DataGridViewRow In Me.dgvItem.Rows
                If IsNothing(rowDGV.Cells(Me.ColIdentifier.Index).Value) = False Then
                    IdentifierD = rowDGV.Cells(Me.ColIdentifier.Index).Value.ToString.IndexOf("-", 1)
                    IdentifierD = rowDGV.Cells(Me.ColIdentifier.Index).Value.ToString.Substring(IdentifierD + 1)
                    If NilaiMax < IdentifierD Then
                        NilaiMax = IdentifierD
                    End If
                End If
            Next
            NilaiMax += 1

            Retval = Me.TxtPrimaryPurchaseOrder.Text + "-" + NilaiMax.ToString
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub dgvItem_CellBeginEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'JIka ada IPO maka Qty Tidak Bisa di edit
            If e.ColumnIndex = Me.ColQty.Index And Me.dgvItem.Item(Me.ColIPONo.Index, e.RowIndex).Value <> "" Then
                e.Cancel = True
                Exit Sub
            End If

            'yang bisa diedit adalah:
            '1. Qty
            '2. Unit Price
            '3. Diskon
            '4. Delivery Date Controll
            '6. Remarks

            ''@rosadinnn
            ''@MUaT
            ''@update kolom yg bisa diedit
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = Me.ColUnitPrice.Index Or e.ColumnIndex = Me.ColDiscount.Index _
                Or e.ColumnIndex = Me.ColDeliveryCont.Index Or e.ColumnIndex = Me.ColRemark.Index Or e.ColumnIndex = Me.ColDiscAmount.Index _
                Or e.ColumnIndex = Me.ColQtyMulti.Index _
                Or e.ColumnIndex = Me.ColToleranceDown.Index Or e.ColumnIndex = Me.ColToleranceUp.Index _
                Or e.ColumnIndex = Me.ColToleranceDownPerc.Index Or e.ColumnIndex = Me.ColToleranceUpPerc.Index Then
                    'kalo mau edit delivery date, maka harus memiliki hak akses-nya
                    If e.ColumnIndex = Me.ColDeliveryCont.Index Then
                        If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiPO_Override_Promise_To_Deliver") = False Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If

                    'kalo sampe sini maka boleh edit
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    'buang nol di belakang decimal jika edit kolom qty

                    ''@rosadinnn
                    ''@MUaT
                    ''@update tambah kolom qty multi
                    If e.ColumnIndex = Me.ColQty.Index Or _
                    e.ColumnIndex = Me.ColDiscAmount.Index Or _
                    e.ColumnIndex = Me.ColDiscount.Index Or _
                    e.ColumnIndex = Me.ColUnitPrice.Index Or _
                    e.ColumnIndex = Me.ColQtyMulti.Index Or _
                    e.ColumnIndex = Me.ColRatioDefault.Index Or _
                    e.ColumnIndex = Me.ColRatioReference.Index Or _
                    e.ColumnIndex = Me.ColToleranceDown.Index Or _
                    e.ColumnIndex = Me.ColToleranceUp.Index Or _
                    e.ColumnIndex = Me.ColToleranceDownPerc.Index Or _
                    e.ColumnIndex = Me.ColToleranceUpPerc.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                'hanya bisa edit pada saat detail aktif
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    ''@rosadinnn
    ''@MUaT
    ''@pergeseran kolom qty ke qty multi - > cek ya
    Private Sub EndEditQty_Multi(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
      ByVal row As DataGridViewRow)
        Dim JmlKurang As Decimal = 0

        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
                'digit grouping

                ''@rosadinnn
                ''@MUaT
                ''cek pob multi, jangan sampai POB minus
                Dim TotalPackingList, TotalStockInPurchase, TotalPOC As Decimal
                Dim msg As String = ""
                Dim productDefPrec As Decimal
                TotalPOC = Me.ObjPO.GetTotalPOC(TxtPrimaryPurchaseOrder.Text, .Cells(ColPrimaryCodeNo.Index).Value)
                TotalStockInPurchase = Me.ObjPO.GetTotalStockInPurchase(TxtPrimaryPurchaseOrder.Text, .Cells(ColPrimaryCodeNo.Index).Value)
                TotalPackingList = Me.ObjPO.GetTotalPackingList(TxtPrimaryPurchaseOrder.Text, .Cells(ColPrimaryCodeNo.Index).Value)
                If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) - (TotalPOC + TotalStockInPurchase + TotalPackingList) < 0 Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    productDefPrec = ObjProduct.GetAngkaDecimal(.Cells(ColPrimaryCodeNo.Index).Value, .Cells(ColPrimaryUnitMulti.Index).Value)
                    If TotalPOC > 0 Then
                        msg = "POC: " & ObjCD.DigitGrouping(TotalPOC, productDefPrec)
                    End If
                    If TotalStockInPurchase > 0 Then
                        If msg = "" Then
                            msg = "Stock in Purchase: " & ObjCD.DigitGrouping(TotalStockInPurchase, productDefPrec)
                        Else
                            msg += vbNewLine & "Stock in Purchase: " & ObjCD.DigitGrouping(TotalStockInPurchase, productDefPrec)
                        End If
                    End If
                    If TotalPackingList > 0 Then
                        If msg = "" Then
                            msg = "Packing List: " & ObjCD.DigitGrouping(TotalPackingList, productDefPrec)
                        Else
                            msg += vbNewLine & "Packing List: " & ObjCD.DigitGrouping(TotalPackingList, productDefPrec)
                        End If
                    End If
                    MsgBox("Minimum value can enter is: " & ObjCD.DigitGrouping((TotalPOC + TotalStockInPurchase + TotalPackingList), productDefPrec) _
                                & ", because there's already: " & vbNewLine & msg, MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                ''update source defprec
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, _
                            ObjProduct.GetAngkaDecimal(Me.dgvItem(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value))
                Me.SetColQtyDefault(row)

                If CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value) = 0 Then
                    MsgBox("Qty restored to previous value, because qty you entered make default qty=0.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.SetColQtyDefault(row)
                    Exit Sub
                End If

                Me.SetTolerancePerc(row.Index)

                'hitung ulang row
                Me.Hitung_Amount(row)
                'hitung total
                Me.Hitung_Total()

                If Me.dgvItem.Item(Me.ColDeliveryCont.Index, e.RowIndex).Value = Boolean.TrueString Then
                    If CDec(Me.Buffer) > CDec(Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value) Then
                        'untuk item delivery, jika qty turun, maka hapuskan mulai dari baris paling bawah
                        JmlKurang = CDec(Me.Buffer) - CDec(Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value)
                        Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value + "'"
                        Me.vTBLItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                        While JmlKurang > 0
                            If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                                Me.vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(0).Row

                                ''@rosadinnn
                                ''@MUaT
                                ''@ganti kolom qty dg yg multi
                                If Me.vrowItemDelivery.delivery_qty_multi > JmlKurang Then
                                    Me.vrowItemDelivery.delivery_qty_multi -= JmlKurang

                                    Me.vrowItemDelivery.delivery_qty = Me.ObjCD.DigitGrouping(Me.GetQtyDefault(Me.vrowItemDelivery.delivery_qty_multi, _
                                            row.Cells(Me.ColRatioDefault.Index).Value, row.Cells(Me.ColRatioReference.Index).Value), _
                                            Me.ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value))

                                    JmlKurang = 0
                                Else

                                    JmlKurang -= Me.vrowItemDelivery.delivery_qty_multi
                                    Me.vrowItemDelivery.delivery_qty_multi = 0

                                End If

                                If Me.vrowItemDelivery.delivery_qty_multi = 0 Then
                                    Me.vTBLItemDelivery.Rows.Remove(Me.vrowItemDelivery)
                                End If

                            End If
                        End While
                    ElseIf CDec(Me.Buffer) < CDec(Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value) Then
                        'untuk item delivery, jika qty naik, maka tambahkan sisanya pada baris paling bawah
                        JmlKurang = CDec(Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value) - CDec(Me.Buffer)
                        Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value + "'"
                        Me.vTBLItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                        If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                            Me.vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(0).Row

                            ''@rosadinnn
                            ''@MUaT
                            Me.vrowItemDelivery.delivery_qty_multi += JmlKurang
                            Me.vrowItemDelivery.delivery_qty = Me.ObjCD.DigitGrouping(Me.GetQtyDefault(Me.vrowItemDelivery.delivery_qty_multi, _
                                                            row.Cells(Me.ColRatioDefault.Index).Value, row.Cells(Me.ColRatioReference.Index).Value), _
                                                            Me.ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value))


                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' fungsi ini digunakan untuk menyamakan semua nilai di dgv ketika terjadi perubahan data
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SamakanDelDateCont_UnitPrice_DiscountDGVItem(ByVal vProductCode As String, _
    ByVal vDelDateCont As Boolean, _
    ByVal vUnitPrice As Decimal, ByVal vDiscount As Decimal, ByVal vDeliveryDate As String, ByVal vNettPrice As Decimal)
        Dim row As DataGridViewRow
        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColCodeNo.Index).Value = vProductCode Then
                    row.Cells(Me.ColDeliveryCont.Index).Value = vDelDateCont
                    row.Cells(Me.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(vUnitPrice, Me.DefaultPrecision))
                    row.Cells(Me.ColDiscount.Index).Value = ObjCD.DigitGrouping(vDiscount)
                    row.Cells(Me.ColDeliveryDate.Index).Value = vDeliveryDate
                    row.Cells(Me.ColNettPrice.Index).Value = ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(vNettPrice, Me.DefaultPrecision))
                    Me.Hitung_Amount(row)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    ''@rosadinnn
    ''@MUaT
    ''@pergeseran kolom qty ke qty multi - > cek ya
    Private Sub EndEditUnitPrice(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai unit price, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for unit price.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value)
                'jika unit price < nett price maka kembalikan ke nilai buffer
                'hanya berlaku klo menentukan diskon berdasarkan amount
                If Me.dgvItem.Item(Me.ColDiscFromPerct.Index, e.RowIndex).Value = "0" Then
                    If CDec(Me.dgvItem.Item(Me.ColDiscAmount.Index, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
                'Jika ada yang sama maka harus di samakan
                Me.SamakanDelDateCont_UnitPrice_DiscountDGVItem(row.Cells(Me.ColCodeNo.Index).Value, _
                     row.Cells(Me.ColDeliveryCont.Index).Value, row.Cells(Me.ColUnitPrice.Index).Value, _
                     row.Cells(Me.ColDiscount.Index).Value, row.Cells(Me.ColDeliveryDate.Index).Value, _
                     row.Cells(Me.ColNettPrice.Index).Value)
                'hitung ulang row
                Me.Hitung_Amount(row)
                'hitung total
                Me.Hitung_Total()
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditDiskonAmount(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai disc amount, pastikan nilainya positif
                If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for discount amount.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, DefaultPrecision)

                If CDec(Me.Buffer) <> CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    'jika unit price < nett price maka kembalikan ke nilai buffer
                    'hanya berlaku klo menentukan diskon berdasarkan amount
                    If CDec(Me.dgvItem.Item(Me.ColDiscAmount.Index, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'tandai diskon dihitung berdasarkan amount
                    row.Cells(Me.ColDiscFromPerct.Index).Value = "0"
                    'hitung ulang row
                    Me.Hitung_Amount(row)
                    'hitung total
                    Me.Hitung_Total()
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditDiscount(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai diskon, pastikan nilainya tidak negatif
                If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for discount.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
                'cek nili maksimum untuk diskon adalah 100%
                If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                    MsgBox("Maximum value is 100%.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, 2)
                'hanya hitung ulang kalo ada perubahan data
                If CDec(Me.Buffer) <> CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    'tandai di ubah dari diskon
                    row.Cells(Me.ColDiscFromPerct.Index).Value = "1"
                    'Jika ada yang sama maka harus di samakan
                    Me.SamakanDelDateCont_UnitPrice_DiscountDGVItem(row.Cells(Me.ColCodeNo.Index).Value, _
                         row.Cells(Me.ColDeliveryCont.Index).Value, row.Cells(Me.ColUnitPrice.Index).Value, _
                         row.Cells(Me.ColDiscount.Index).Value, row.Cells(Me.ColDeliveryDate.Index).Value, _
                         row.Cells(Me.ColNettPrice.Index).Value)
                    'hitung total
                    Me.Hitung_Total()
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPreviousRowDeliveryDate(ByVal vPreviousRowIndex As Integer) As Date
        Try
            If vPreviousRowIndex > -1 Then
                If Me.dgvItem.Item(Me.ColDeliveryCont.Index, vPreviousRowIndex).Value = True Then
                    If IsDate(Me.dgvItem.Item(Me.ColDeliveryDate.Index, vPreviousRowIndex).Value) = True Then
                        Return CDate(Me.dgvItem.Item(Me.ColDeliveryDate.Index, vPreviousRowIndex).Value)
                    Else
                        Me.vTBLItemDelivery.DefaultView.RowFilter = "Identifier = '" & Me.dgvItem.Item(Me.ColIdentifier.Index, vPreviousRowIndex).Value & "'"
                        If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                            Me.vTBLItemDelivery.DefaultView.Sort = "[delivery date] ASC"
                            Me.vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(0).Row
                            Return Me.vrowItemDelivery.delivery_date
                        End If
                    End If
                End If
            End If
            Return Me.dtpDate.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Me.dtpDate.Value
        End Try
    End Function

    Private Sub EndEditDeliveryDateCont(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
      ByVal row As DataGridViewRow)
        Dim rowD As DseTransaksiPO.tbltrans_po_itemdeliveryRow
        Dim ij As Integer = 0

        Try
            With row
                If row.Cells(Me.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                    'jika jadi true maka delivery date jadi tanggal transaksi
                    row.Cells(Me.ColDeliveryDate.Index).Value = Format(Me.GetPreviousRowDeliveryDate(e.RowIndex - 1), "dd MMM yyyy")

                    'tambahkan ke dalam item delivery date
                    rowD = Me.vTBLItemDelivery.Newtbltrans_po_itemdeliveryRow
                    With rowD
                        '.delivery_date = Me.dtpDate.Value
                        .delivery_date = Me.GetPreviousRowDeliveryDate(e.RowIndex - 1)
                        .delay = Me.ObjItem.GetPreparationDelay(row.Cells(Me.ColPrimaryCodeNo.Index).Value)
                        .delivery_qty = Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value
                        .identifier = Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value
                        .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                        .primary_transaksi = Me.TxtPrimaryPurchaseOrder.Text
                        .remarks = ""

                        ''@rosadinnn
                        ''@MUaT
                        .delivery_qty_multi = Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value

                    End With
                    Me.vTBLItemDelivery.Rows.Add(rowD)
                Else
                    'jika jadi false maka delivery date jadi "-"
                    row.Cells(Me.ColDeliveryDate.Index).Value = "-"
                    'buang item delivery
                    While ij < Me.vTBLItemDelivery.Rows.Count
                        Me.vrowItemDelivery = Me.vTBLItemDelivery.Rows(ij)
                        If Me.vrowItemDelivery.identifier = Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value Then
                            Me.vTBLItemDelivery.Rows.Remove(Me.vrowItemDelivery)
                        Else
                            ij += 1
                        End If
                    End While
                End If
                Me.SamakanDelDateCont_UnitPrice_DiscountDGVItem(row.Cells(Me.ColCodeNo.Index).Value, _
                row.Cells(Me.ColDeliveryCont.Index).Value, row.Cells(Me.ColUnitPrice.Index).Value, _
                row.Cells(Me.ColDiscount.Index).Value, row.Cells(Me.ColDeliveryDate.Index).Value, _
                row.Cells(Me.ColNettPrice.Index).Value)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim ij As New FrmTransaksiPO_SetDelivery
        Dim DI As Integer = 0
        Dim rowD As DseTransaksiPO.tbltrans_po_itemdeliveryRow
        Dim JmlDelivery As Integer = 0
        Dim DelDateD As Date = Today

        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColDeliveryDate.Index Then
                If Me.dgvItem.Item(Me.ColDeliveryDate.Index, e.RowIndex).Value <> "-" And _
                Me.dgvItem.Item(Me.ColDeliveryCont.Index, e.RowIndex).Value = Boolean.TrueString Then
                    ij.vPrimaryCodeNo = Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value

                    ''@rosadinnn
                    ''@MUaT
                    ij.vPrimaryUnitMulti = Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value
                    ij.vRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                    ij.vRatioReference = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                    ij.vPrimaryUnitDefault = Me.dgvItem.Item(Me.ColPrimaryUnits.Index, e.RowIndex).Value

                    ij.VIdentifier = Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value
                    ij.FrmRef = Me
                    If Me.TxtRemark.Enabled = False Then
                        ij.vReadOnly = True
                        ij.ShowDialog()
                        Exit Sub
                    End If
                    If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                        'hapuskan semua item delivery yang identifiernya di edit
                        DI = 0
                        While DI <= Me.vTBLItemDelivery.Rows.Count - 1
                            Me.vrowItemDelivery = Me.vTBLItemDelivery.Rows(DI)
                            If Me.vrowItemDelivery.identifier = ij.VIdentifier Then
                                Me.vTBLItemDelivery.Rows.Remove(Me.vrowItemDelivery)
                            Else
                                DI += 1
                            End If
                        End While
                        'masukkan data item delivery baru
                        ij.dgvDetail.AllowUserToAddRows = False
                        For Each rowDGV As DataGridViewRow In ij.dgvDetail.Rows
                            rowD = Me.vTBLItemDelivery.Newtbltrans_po_itemdeliveryRow
                            With rowD
                                .delivery_date = CDate(rowDGV.Cells(ij.ColDeliveryDate.Index).Value)
                                .delay = rowDGV.Cells(ij.ColDelay.Index).Value
                                .delivery_qty = rowDGV.Cells(ij.ColQty.Index).Value
                                .identifier = ij.VIdentifier
                                .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                                .primary_transaksi = Me.TxtPrimaryPurchaseOrder.Text
                                .remarks = rowDGV.Cells(ij.ColRemarks.Index).Value

                                ''@rosadinnn
                                ''@MUaT
                                .delivery_qty_multi = rowDGV.Cells(ij.ColQtyMulti.Index).Value

                            End With
                            Me.vTBLItemDelivery.Rows.Add(rowD)
                            JmlDelivery += 1
                            DelDateD = CDate(rowDGV.Cells(ij.ColDeliveryDate.Index).Value)
                        Next

                        If JmlDelivery > 1 Then
                            Me.dgvItem.Item(Me.ColDeliveryDate.Index, e.RowIndex).Value = "Multiple"
                        ElseIf JmlDelivery = 1 Then
                            Me.dgvItem.Item(Me.ColDeliveryDate.Index, e.RowIndex).Value = Format(DelDateD, "dd MMM yyyy")
                        End If
                    End If
                End If

                If e.RowIndex = Me.dgvItem.RowCount - 1 Then
                    If Me.dgvItem.Rows(e.RowIndex).Cells(Me.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                        '  Me.dtpDeliveryDate.Checked = True
                        If IsDate(Me.dgvItem.Rows(e.RowIndex).Cells(Me.ColDeliveryDate.Index).Value) Then
                            Me.dtpDeliveryDate.Value = Me.dgvItem.Rows(e.RowIndex).Cells(Me.ColDeliveryDate.Index).Value
                        End If
                    End If
                End If
            End If

            ''@rosadinnn
            ''@MUaT
            Dim frmRatio As New FrmTransaksiPO_SetQtyRatio
            Dim row As DataGridViewRow
            If e.ColumnIndex = Me.ColRatioDesc.Index Then
                If Me.dgvItem.Item(Me.ColPrimaryUnits.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value Or _
                    Me.TxtRemark.Enabled = False Then
                    Exit Sub
                End If
                Dim BufferRatioReference As String = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                Dim BufferRatioDefault As String = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                frmRatio.vDefPrecMulti = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value)
                frmRatio.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                frmRatio.vRatioMulti = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                frmRatio.vUnitMulti = Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value
                frmRatio.vRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                frmRatio.vUnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                If frmRatio.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = frmRatio.TxtRatioMulti.Text
                    Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = frmRatio.TxtRatioDefault.Text
                    row = Me.dgvItem.Rows(e.RowIndex)
                    Me.SetRatioAndQtyDefault(row, True)

                    If CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value) <= 0 Then
                        MsgBox("Ratio value restored to previous value, because ratio you entered make default qty=0.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                        Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                        Me.SetRatioAndQtyDefault(row, True)
                    Else
                        Me.SetTolerancePerc(row.Index)
                        Me.vTBLItemDelivery.DefaultView.RowFilter = "[identifier] ='" & dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value & "'"
                        If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                            For i As Integer = 0 To Me.vTBLItemDelivery.DefaultView.Count - 1
                                Me.vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(i).Row
                                Me.vrowItemDelivery.delivery_qty = ObjCD.DigitGrouping(vrowItemDelivery.delivery_qty_multi * CDec(frmRatio.TxtRatioDefault.Text) / CDec(frmRatio.TxtRatioMulti.Text), frmRatio.vDefPrecDefault)
                            Next
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek

                    ''@rosadinnn
                    ''@MUaT
                    ''@update kolom qty multi
                    If e.ColumnIndex = Me.ColQty.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Or _
                    e.ColumnIndex = Me.ColDiscount.Index Or e.ColumnIndex = Me.ColNettPrice.Index Or _
                    e.ColumnIndex = Me.ColAmount.Index Or _
                    e.ColumnIndex = Me.ColQtyMulti.Index Or _
                    e.ColumnIndex = Me.ColRatioDefault.Index Or _
                    e.ColumnIndex = Me.ColRatioReference.Index Or _
                    e.ColumnIndex = Me.ColToleranceDown.Index Or _
                    e.ColumnIndex = Me.ColToleranceUp.Index Or _
                    e.ColumnIndex = Me.ColToleranceDownPerc.Index Or _
                    e.ColumnIndex = Me.ColToleranceUpPerc.Index Then
                        If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            If IsNumeric(Me.Buffer) = True Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Else
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = 0
                            End If
                        End If
                    End If
                    If e.ColumnIndex <> Me.ColDeliveryDate.Index Then
                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                            Exit Sub
                        End If
                    Else
                        If IsDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Date = CDate(Me.Buffer).Date Then
                                Exit Sub
                            End If
                        End If
                    End If

                    If e.ColumnIndex <> Me.ColRemark.Index Then
                        'kalo bukan data remarks maka perlu dicek kebenaran data hasil edit

                        ''@rosadinnn
                        ''@MUaT
                        ''@update kolom yg dicek
                        If e.ColumnIndex = Me.ColQtyMulti.Index Then
                            Me.EndEditQty_Multi(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColUnitPrice.Index Then
                            'unit price
                            Me.EndEditUnitPrice(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColDiscount.Index Then
                            'discount
                            Me.EndEditDiscount(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColDeliveryCont.Index Then
                            'kontrol delivery date
                            Me.EndEditDeliveryDateCont(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColDiscAmount.Index Then
                            Me.EndEditDiskonAmount(sender, e, row)

                        ElseIf e.ColumnIndex = Me.ColRatioDefault.Index Or e.ColumnIndex = Me.ColRatioReference.Index Then
                            ''@harus positif
                            If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Please enter positive number for ratio data.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                            If e.ColumnIndex = Me.ColRatioDefault.Index Then
                                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, _
                                                                            Me.ObjProduct.GetAngkaDecimal(.Cells(Me.ColPrimaryCodeNo.Index).Value))
                            Else
                                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, _
                                ObjProduct.GetAngkaDecimal(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryUnitMulti.Index).Value))
                            End If
                            Me.SetRatioAndQtyDefault(row, True)
                        ElseIf e.ColumnIndex = Me.ColToleranceDown.Index Or e.ColumnIndex = Me.ColToleranceUp.Index Then
                            ''@harus positif
                            If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Please enter positive number for tolerance data.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                            ''maksimal 50% dari qty multi
                            If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value > (Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value / 2) Then
                                MsgBox("Tolerance cannot greater than 50% of qty.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                            Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, _
                                        ObjProduct.GetAngkaDecimal(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryUnitMulti.Index).Value))
                            SetTolerancePerc(e.RowIndex)
                        ElseIf e.ColumnIndex = Me.ColToleranceDownPerc.Index Or e.ColumnIndex = Me.ColToleranceUpPerc.Index Then
                            ''@harus positif
                            If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Please enter positive number for tolerance data.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                            ''maksimal 50%
                            If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value > 50 Then
                                MsgBox("Tolerance cannot greater than 50%.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                            SetToleranceVal(e.RowIndex)
                        End If

                    End If
                Else
                    If e.ColumnIndex = Me.ColRemark.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Me.dgvItem.TabStop = True
        If Me.TxtRemark.Enabled = True Then
            Me.btnUP.Enabled = True
            Me.btnDown.Enabled = True
        End If
    End Sub

    Dim skipHitungTotal As Boolean = False

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        If Me.dgvItem.Rows.Count = 0 Then
            Me.dgvItem.TabStop = False
            If Me.TxtRemark.Enabled = True Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
        End If

        If Me.skipHitungTotal = False Then
            Me.Hitung_Total()
        Else
            Me.skipHitungTotal = False
        End If

    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> 0 Then
                    Me.skipHitungTotal = True
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    If row.Displayed = False Then
                        Me.ShowRow(row, True)
                    End If
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> Me.dgvItem.Rows.Count - 1 Then
                    Me.skipHitungTotal = True
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    Me.ShowRow(row, False)
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRow(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvItem.Rows(row.Index).Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False) Then
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.dgvItem.SelectedRows.Count > 0 Then
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                Else
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Me.ObjCD.SetNumber(Me.dgvItem)
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
            If Me.IPONo <> "" Then
                Me.LoadDGVIPODanPO(Me.TxtPrimaryPurchaseOrder.Text)
            End If

            ''@rosadinnn
            ''@MUaT
            Me.FormatDgv()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim ListIdentifier As New List(Of String)
        Dim rowDGV1 As DataGridViewRow
        Dim Qty, QtyDelivery As Decimal
        Dim CodeNo As String = ""
        Dim Dummy As String = ""

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("PO number")
            KoleksiTextBox.Add(Me.TxtPONo)
            KoleksiNamaData.Add("supplier code")
            KoleksiTextBox.Add(Me.txtSupplierCode)
            KoleksiNamaData.Add("currency")
            KoleksiTextBox.Add(Me.txtCurrencyCode)

            If ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'kalo sampe cek berkaitan dengan timestamp apakah masih benar
            If Me.baru = False Then
                If Me.ObjPO.TimeStamp <> Me.ObjPO.GetTimestamp(Me.ObjPO.primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjPO.CariDetail(Me.ObjPO.primary, Me.dgvItem, Me.vTBLItemDelivery, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'cek berkaitan dengan perubahan tanggal, jika auto number
            If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                If Me.baru = False Then
                    If Me.ObjPO.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                    Me.ObjPO.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika local/impor dari supplier tidak berubah dan tanggal tidak berubah bulan(masih pada bulan yang sama) maka
                        'kembalikan no transaksi ke yang lama
                        If Me.ObjSupplier.ApakahLokal(Me.txtPrimarySupplier.Text) = Me.ObjSupplier.ApakahLokal(Me.ObjPO.primarySupplier) Then
                            Me.TxtPONo.Text = Me.ObjPO.NoTransaksi
                        End If
                    Else
                        If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                            MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                            Me.dtpDate.Value = Me.ObjPO.TanggalTransaksi
                            Me.TxtPONo.Text = Me.ObjPO.NoTransaksi
                            Return False
                        Else
                            Me.TxtPONo.Text = Me.ObjPO.GetNoBaru1(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                        End If
                    End If
                End If
            End If

            ''@rosadinnn
            ''@MUaT
            ''salah jika toleransi melebihi 50% dari qty
            Dim msg_ As String = ""
            For Each rowDgv_ As DataGridViewRow In Me.dgvItem.Rows
                With rowDgv_
                    If CDec(.Cells(Me.ColToleranceDown.Index).Value) > (CDec(.Cells(Me.ColQtyMulti.Index).Value) / 2) Or _
                            CDec(.Cells(Me.ColToleranceUp.Index).Value) > (CDec(.Cells(Me.ColQtyMulti.Index).Value) / 2) Or _
                            CDec(.Cells(Me.ColToleranceDownPerc.Index).Value) > 50 Or _
                            CDec(.Cells(Me.ColToleranceUpPerc.Index).Value) > 50 Then
                        msg_ += .Cells(Me.ColCodeNo.Index).Value & ","
                    End If
                End With
            Next
            If String.IsNullOrEmpty(msg_) = False Then
                msg_ = msg_.Remove(msg_.Length - 1, 1)
                MsgBox("Tolerance value(down or up) of the following product code cannot greater than 50%: " & vbNewLine & msg_, MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan setiap item delivery, qty-nya (multi dan default) konsisten dengan item po
            ListIdentifier.Clear()
            For Each rowDGV1 In Me.dgvItem.Rows
                If rowDGV1.Cells(Me.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                    If ListIdentifier.Contains(rowDGV1.Cells(Me.ColIdentifier.Index).Value) = False Then
                        ListIdentifier.Add(rowDGV1.Cells(Me.ColIdentifier.Index).Value)
                    End If
                End If
            Next
            For Each Me.vrowItemDelivery In Me.vTBLItemDelivery.Rows
                If ListIdentifier.Contains(Me.vrowItemDelivery.identifier) = False Then
                    ListIdentifier.Contains(Me.vrowItemDelivery.identifier)
                End If
            Next
            For Each Identifier As String In ListIdentifier
                'cek qty default
                Qty = Me.GetJmlDGV_Default(Identifier)
                QtyDelivery = Me.GetJmlItemDelivery_Default(Identifier)
                If Qty <> QtyDelivery Then
                    CodeNo = Me.GetCodeNo(Identifier)
                    If CodeNo <> "" Then
                        MsgBox("Qty delivery for: " + CodeNo + " is not consistent, please check.", MsgBoxStyle.Information, "Information")
                    Else
                        MsgBox("Logic error: inconsistent delivery qty, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    End If

                    Return False
                End If

                'cek qty multi
                Qty = Me.GetJmlDGV_Multi(Identifier)
                QtyDelivery = Me.GetJmlItemDelivery_Multi(Identifier)
                If Qty <> QtyDelivery Then
                    CodeNo = Me.GetCodeNo(Identifier)
                    If CodeNo <> "" Then
                        MsgBox("Qty delivery for: " + CodeNo + " is not consistent, please check.", MsgBoxStyle.Information, "Information")
                    Else
                        MsgBox("Logic error: inconsistent delivery qty, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    End If

                    Return False
                End If
            Next

            If Me.chkTax.Checked = True Then
                If CDec(Me.txtTax.Text) <= 0 Then
                    Me.txtTax.Focus()

                    MsgBox("Please specify tax percentage.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                If Dummy = "" Then
                    Dummy = "0"
                End If

                If CDec(Me.txtTax.Text) <> CDec(Dummy) Then
                    Me.txtTax.Focus()

                    If MsgBox("Tax rate <> " + Me.ObjCD.DigitGrouping(Dummy, 2) + "%, continue to save data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Return False
                    End If
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            ListIdentifier = Nothing
        End Try
    End Function

    Private Function GetCodeNo(ByVal vIdentifier As String) As String
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColIdentifier.Index).Value = vIdentifier Then
                    Return rowDGV.Cells(Me.ColCodeNo.Index).Value
                End If
            Next
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetJmlDGV_Default(ByVal vIdentifier As String) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColIdentifier.Index).Value = vIdentifier Then
                    RetVal += CDec(rowDGV.Cells(Me.ColQty.Index).Value)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV_Multi(ByVal vIdentifier As String) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColIdentifier.Index).Value = vIdentifier Then
                    RetVal += CDec(rowDGV.Cells(Me.ColQtyMulti.Index).Value)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlItemDelivery_Default(ByVal vIdentifier As String) As Decimal
        Dim DV As Data.DataRowView
        Dim RetVal As Decimal = 0

        Try
            Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + vIdentifier + "'"
            For Each DV In Me.vTBLItemDelivery.DefaultView
                Me.vrowItemDelivery = DV.Row
                RetVal += Me.vrowItemDelivery.delivery_qty
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlItemDelivery_Multi(ByVal vIdentifier As String) As Decimal
        Dim DV As Data.DataRowView
        Dim RetVal As Decimal = 0

        Try
            Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + vIdentifier + "'"
            For Each DV In Me.vTBLItemDelivery.DefaultView
                Me.vrowItemDelivery = DV.Row
                RetVal += Me.vrowItemDelivery.delivery_qty_multi
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeField()
        Dim DV As Data.DataRowView
        Dim row As Core_Module.DseMasterCompanyData.tbltrans_company_data_itemRow
        Dim i As Integer = 0
        Dim SupName, CPName, PhoneNo As String

        Try
            SupName = ""
            CPName = ""
            PhoneNo = ""

            With Me.ObjPO
                Me.txtContactPerson.Text = .SuppContactPerson
                Me.dtpDate.Value = .TanggalTransaksi
                Me.TxtPrimaryPurchaseOrder.Text = .primary
                '        Me.TxtPrimaryQuotation.Text = .primary
                Me.TxtPONo.Text = .NoTransaksi
                Me.txtPrimarySupplier.Text = .primarySupplier
                Me.txtSupplierAddIdentifier.Text = .IdentifierSuppAddress
                Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                ObjSupplier.Get_SuppName_ContactPerson_PhoneNo(Me.txtPrimarySupplier.Text, SupName, CPName, PhoneNo)
                Me.TxtSupplierName.Text = SupName
                Me.ObjSupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                Me.cboSupplierAddress.Text = Me.ObjSupplier.GetKodeAddress(Me.txtSupplierAddIdentifier.Text)
                Me.txtContactPerson.Text = .SuppContactPerson
                Me.txtPhoneNo.Text = PhoneNo
                Me.TxtShipVia.Text = .ShipVia
                Me.txtPrimaryCurrency.Text = .PrimaryCurrency
                Me.txtCurrencyCode.Text = Me.objCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                Me.txtPaymentTerm.Text = .PaymentTerm
                Me.txtPaymentDueDate.Text = .PaymentDue
                Me.SetDefaultPrecision()
                Me.cboPaymentDueDateReference.SelectedIndex = .ReferencePaymentDue
                Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubTotal, Me.DefaultPrecision))
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.TaxAmount, Me.DefaultPrecision))
                Me.txtTax.Text = .Tax
                '    Me.txtTermsNCondition.Text = .TermCondition
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Total, Me.DefaultPrecision))
                Me.chkTax.Checked = .TaxExclude
                Me.TxtRemark.Text = .Remarks
                If Me.ObjPO.Authorizationlevel = 0 Then
                    Me.txtAutohirizationStatus.Text = "Unauthorized"
                ElseIf Me.ObjPO.Authorizationlevel = 1 Then
                    Me.txtAutohirizationStatus.Text = "Authorized By Purchase Staff"
                ElseIf Me.ObjPO.Authorizationlevel = 2 Then
                    Me.txtAutohirizationStatus.Text = "Authorized By Purchase Manager"
                ElseIf Me.ObjPO.Authorizationlevel = 3 Then
                    Me.txtAutohirizationStatus.Text = "Authorized By Director"
                End If
                Me.TxtIdentifierShipTo.Text = .IdentifierShipTo
                For Each DV In Me.TBL_Add.DefaultView
                    row = DV.Row
                    If row.Add_Identifier = Me.TxtIdentifierShipTo.Text Then
                        Me.cboShipTo.SelectedIndex = i
                        Exit For
                    End If
                    i += 1
                Next

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                ''@rosadinnn
                ''@MUaT
                Me.FormatDgv()

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.ObjPO
                If Me.baru = True Then
                    .Authorizationlevel = 0
                    .AuthorizationNeeded = 0
                    .PurchaseDirectorAuthorizationDate = #1/1/1950#
                    .PrimaryDirectorStaffAuthorization = -1
                    .PurchaseManagerAuthorizationDate = #1/1/1950#
                    .PrimaryManagerStaffAuthorization = -1
                    .PurchaseStaffAuthorizationDate = #1/1/1950#
                    .PrimaryPurchaseStaffAuthorization = -1
                End If
                .SuppContactPerson = Me.txtContactPerson.Text
                .ShipVia = Me.TxtShipVia.Text
                .SuppPhoneNo = Me.txtPhoneNo.Text
                .IdentifierSuppAddress = Me.txtSupplierAddIdentifier.Text
                .NoTransaksi = Me.TxtPONo.Text
                .PaymentDue = Me.txtPaymentDueDate.Text
                .PaymentTerm = Me.txtPaymentTerm.Text
                .PrimaryCurrency = Me.txtPrimaryCurrency.Text
                .primarySupplier = Me.txtPrimarySupplier.Text
                .ReferencePaymentDue = Me.cboPaymentDueDateReference.SelectedIndex
                .SubTotal = Me.txtSubTotal.Text
                .TaxAmount = Me.txtTaxAmount.Text
                .Tax = Me.txtTax.Text
                .Total = Me.TxtTotal.Text
                .TaxExclude = Me.chkTax.Checked
                .Remarks = Me.TxtRemark.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .IdentifierShipTo = Me.TxtIdentifierShipTo.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.KendalikanControl(SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)

            Me.ObjCD.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            'Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.chkInsertFromIPO.Checked = False
            Me.llbHistoricalPrice.Enabled = False

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(SplitContainer2.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(SplitContainer2.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(SplitContainer3.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(SplitContainer3.Panel2.Controls, Me.BackColor, False)
            Me.llbUpdateDeliveryDate.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            'pastikan numbering purchase order lokal dan numbering purchase order impor sama nilai autonumbernya
            If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor) <> Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) Then
                MsgBox("Please make sure numbering system for local and import PO are identical.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(SplitContainer2.Panel1.Controls, Me.TxtPrimaryPurchaseOrder.BackColor, False)
            Me.ObjCD.SetBackColor(SplitContainer2.Panel2.Controls, Me.TxtPrimaryPurchaseOrder.BackColor, False)
            Me.ObjCD.SetBackColor(SplitContainer3.Panel1.Controls, Me.TxtPrimaryPurchaseOrder.BackColor, False)
            Me.ObjCD.SetBackColor(SplitContainer3.Panel2.Controls, Me.TxtPrimaryPurchaseOrder.BackColor, False)

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryPurchaseOrder.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryPurchaseOrder.BackColor, False)

            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.KendalikanControl(SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)

            Me.ObjCD.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If

            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False
            Me.llbHistoricalPrice.Enabled = True

            If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                Me.TxtPONo.ReadOnly = True
                Me.TxtPONo.BackColor = Me.TxtReadOnly.BackColor
                Me.TxtPONo.TabStop = False
                Me.dtpDate.Focus()
            Else
                Me.TxtPONo.ReadOnly = False
                Me.TxtPONo.BackColor = Me.TxtPrimaryPurchaseOrder.BackColor
                Me.TxtPONo.TabStop = True
                Me.TxtPONo.Focus()
            End If

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiPO_Override_Promise_To_Deliver") = False Then
                Me.dtpDeliveryDate.Enabled = False
                Me.dtpDeliveryDate.Checked = False
            Else
                Me.dtpDeliveryDate.Enabled = True
            End If

            ''@rosadinnn
            ''@MUaT
            Me.FormatDgv()
            Me.dtpDeliveryDate.Checked = False
            Me.llbUpdateDeliveryDate.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.KosongkanItem()
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjPO.SimpanBaru(Me.dgvItem, Me.UserCode, Me, Me.vTBLItemDelivery) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.LangsungClose = True
                            Me.DialogResult = Windows.Forms.DialogResult.OK
                            Me.Close()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjPO.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me, Me.vTBLItemDelivery) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.LangsungClose = True
                            Me.DialogResult = Windows.Forms.DialogResult.OK
                            Me.Close()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, False)

            ''@rosadinnn
            ''@MUaT
            ObjCD.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, False)
            ObjCD.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, False)

            Me.TxtPrimaryPurchaseOrder.Text = "-1"
            Me.txtPrimaryCurrency.Text = "-1"
            Me.txtPrimarySupplier.Text = "-1"
            Me.txtPaymentDueDate.Text = "0"
            Me.txtSubTotal.Text = "0"
            Me.txtTax.Text = "0"
            Me.txtTaxAmount.Text = "0"
            Me.TxtTotal.Text = "0"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()
            Me.cboPaymentDueDateReference.SelectedIndex = 0
            Me.chkTax.Checked = True
            Me.cboSupplierAddress.DataSource = Nothing
            Me.cboSupplierAddress.SelectedIndex = -1
            Me.cboSupplierAddress.Items.Clear()
            Me.vTBLItemDelivery.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtRemark.Text = "" And Me.dgvItem.RowCount = 0 And _
            Me.TxtCodeNo.Text = "" And Me.TxtRemarkItem.Text = "" Then
                Me.NonAktifkanDetail()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.KosongkanItem()
                    If baru = True Then
                        Me.KosongkanField()
                        Me.LangsungClose = True
                        Me.Close()
                    Else
                        Me.LangsungClose = True
                        Me.Close()
                    End If
                Else
                    Exit Sub
                End If
                'Me.NonAktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim ij As New FrmCariTransaksiPO
    '    Try
    '        ij.BackColorAll = Me.BAckColorAll
    '        ij.ShowDialog()
    '        If Not ij.PrimaryPO = -1 Then
    '            If Me.ObjPO.CariDetail(ij.PrimaryPO, Me.dgvItem, Me.vTBLItemDelivery, Me) _
    '            = True Then
    '                Me.SalinKeField()
    '                Me.btnEdit.Focus()
    '            Else
    '                Me.btnFind.Focus()
    '            End If
    '            Me.btnFind.Focus()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ij.Dispose()
    '    End Try
    'End Sub

    'Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
    '            MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
    '            MsgBoxStyle.Information, "Information")
    '            Exit Sub
    '        End If

    '        'pastikan cara penomoran transaksi untuk po lokal dan po impor sama
    '        If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor) <> _
    '        Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) Then
    '            MsgBox("Please specify similar numbering convention for local and import purchase order", MsgBoxStyle.Information, "Information")
    '            Exit Sub
    '        End If

    '        Me.baru = True
    '        Me.KosongkanField()
    '        Me.AktifkanDetail()
    '        Me.ObjPO.ResetObject(Me.dgvItem)
    '        'secara default yang melakukan tanda tangan pada PO adalah, ambil dari setting PO
    '        Me.ObjPO.PreparedBy = My.Settings.PO_Prepared_By
    '        Me.ObjPO.PreparedBy_Jabatan = My.Settings.PO_Prepared_Position
    '        Me.ObjPO.ApprovedBy = My.Settings.PO_Approved_By
    '        Me.ObjPO.ApprovedBy_Jabatan = My.Settings.PO_Approved_Position
    '        Me.ObjPO.AcknowledgeBy = My.Settings.PO_Acknowled_By
    '        Me.ObjPO.AcknowledgeBy_Jabatan = My.Settings.PO_Acknowledge_Position

    '        If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
    '            Me.TxtPONo.Text = Me.ObjPO.GetNoBaru1(Me.dtpDate.Value)
    '        End If
    '        Me.txtAutohirizationStatus.Text = "Unautorized"

    '        Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
    '        Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
    '        Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
    '        Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

    '        'kalo new mata uang langsung IDR
    '        Me.txtPrimaryCurrency.Text = "1"
    '        Me.txtCurrencyCode.Text = Me.objCurrency.GetCurrencyCode(1)
    '        Me.chkTax.Checked = True
    '        Me.txtTax.Text = Me.ObjCD.DigitGrouping("10", 2)

    '        'isikan ulang cbo ship to, dan set index default
    '        Me.ObjCompanyData.IsiCboAdd(Me.TBL_Add, Me.cboShipTo)
    '        Me.SetIndexCboShipToDefault()

    '        'secara default payment term adalah dr invoice date
    '        Me.cboPaymentDueDateReference.SelectedIndex = 1
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        If Me.TxtPONo.Text = "" Then
    '            Me.btnFind.PerformClick()
    '        End If

    '        If Me.TxtPONo.Text <> "" Then
    '            If Me.ObjPO.ApakahBolehEdit(Me.ObjPO.primary, UserCode, True, True) = True Then
    '                Me.baru = False
    '                Me.ObjPO.CariDetail(Me.ObjPO.primary, Me.dgvItem, Me.vTBLItemDelivery, Me)
    '                Me.SalinKeField()
    '                Me.AktifkanDetail()
    '                Me.dgvItem.Focus()
    '                'nilai min juga harus diatas tanggal pengkuncian data terakhir
    '                'msgbox("belum")
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim hasil As MsgBoxResult
    '    Try
    '        If Me.TxtPONo.Text = "" Then
    '            Me.btnFind.PerformClick()
    '        End If

    '        If Me.TxtPONo.Text <> "" Then
    '            If Me.ObjPO.ApakahBolehHapus(Me.TxtPrimaryPurchaseOrder.Text, True, UserCode, True) = False Then
    '                Me.btnFind.Focus()
    '            Else
    '                hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
    '                If hasil = MsgBoxResult.Yes Then
    '                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
    '                    If Me.ObjPO.Hapus(Me.TxtPrimaryPurchaseOrder.Text, UserCode, True) = True Then
    '                        Me.KosongkanField()
    '                        Me.btnNew.Focus()
    '                        MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
    '                    End If
    '                End If
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        Windows.Forms.Cursor.Current = Cursors.Arrow
    '    End Try
    'End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNo.KeyDown
        Dim i As New Core_Module.frmCariMasterProduct
        Dim Jml As Long = 0
        i.BackColorAll = Me.BAckColorAll
        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.TxtCodeNo.Text)
                If Jml = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.dtpDeliveryDate.Focus()
                    Me.SetGbo()

                    ''@rosadinnn
                    ''@MUaT
                    Me.PrepareMultiUnit(Me.TxtPrimaryCodeNo.Text)

                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.dtpDeliveryDate.Focus()
                        Me.SetGbo()

                        ''@rosadinnn
                        ''@MUaT
                        Me.PrepareMultiUnit(Me.TxtPrimaryCodeNo.Text)

                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.Leave
        Dim i As New Core_Module.frmCariMasterProduct
        i.BackColorAll = Me.BAckColorAll
        Try
            If Me.TxtCodeNo.Text <> "" Then
                If Me.ObjProduct.GetJml(Me.TxtCodeNo.Text) = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.dtpDeliveryDate.Focus()
                    Me.SetGbo()

                    ''@rosadinnn
                    ''@MUaT
                    Me.PrepareMultiUnit(Me.TxtPrimaryCodeNo.Text)

                Else
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.dtpDeliveryDate.Focus()
                        Me.SetGbo()

                        ''@rosadinnn
                        ''@MUaT
                        Me.PrepareMultiUnit(Me.TxtPrimaryCodeNo.Text)

                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Clear()
                Me.txtPrimaryCategory.Text = "-1"
                Me.KosongkanItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub LoadDGVIPODanPO(ByVal vPrimaryPO As Double)
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_po_item11TableAdapter
        Dim TBL As New DseTransaksiPO.tbltrans_po_item11DataTable
        Dim row As DseTransaksiPO.tbltrans_po_item11Row

        Try
            If Me.cboFilterIPO.SelectedItem = "Code No" Then
                Me.TxtIPONo.Enabled = True
                Me.btnFindIPO.Enabled = True
                Me.TxtIPONo.CharacterCasing = CharacterCasing.Upper
                ObjPO.IsiDGVIPOByProductCode(Me.dgvItemIPO, Me.TxtIPONo.Text)
                DTA.FillBy_PrimaryPONProductCodeLike(TBL, vPrimaryPO, Me.TxtIPONo.Text + "%")
            ElseIf Me.cboFilterIPO.SelectedItem = "Description" Then
                Me.TxtIPONo.Enabled = True
                Me.btnFindIPO.Enabled = True
                Me.TxtIPONo.CharacterCasing = CharacterCasing.Normal
                ObjPO.IsiDGVIPOByProductDesc(Me.dgvItemIPO, Me.TxtIPONo.Text)
                DTA.FillBy_PrimaryPONDescriptionlike(TBL, vPrimaryPO, Me.TxtIPONo.Text + "%")
            ElseIf Me.cboFilterIPO.SelectedItem = "IPO No" Then
                Me.TxtIPONo.Enabled = True
                Me.btnFindIPO.Enabled = True
                Me.TxtIPONo.CharacterCasing = CharacterCasing.Upper
                ObjPO.IsiDGVIPOByIPONo(Me.dgvItemIPO, Me.TxtIPONo.Text)
                DTA.FillBy_PrimaryPONNoIPOLike(TBL, vPrimaryPO, Me.TxtIPONo.Text + "%")
            ElseIf Me.cboFilterIPO.SelectedItem = "No Filter" Then
                Me.TxtIPONo.Clear()
                Me.TxtIPONo.Enabled = False
                Me.btnFindIPO.Enabled = False
                ObjPO.IsiDGVIPOByAll(Me.dgvItemIPO)
                DTA.FillBy_PrimaryPO(TBL, vPrimaryPO)
            ElseIf Me.cboFilterIPO.SelectedItem = "Supplier Code" Then
                Me.TxtIPONo.Enabled = True
                Me.btnFindIPO.Enabled = True
                Me.TxtIPONo.CharacterCasing = CharacterCasing.Upper
                ObjPO.IsiDGVIPOBySupplierCode(Me.dgvItemIPO, Me.TxtIPONo.Text)
                DTA.FillBy_PrimaryPONSupplierCode(TBL, vPrimaryPO, Me.TxtIPONo.Text + "%")
            ElseIf Me.cboFilterIPO.SelectedItem = "Supplier Name" Then
                Me.TxtIPONo.Enabled = True
                Me.btnFindIPO.Enabled = True
                Me.TxtIPONo.CharacterCasing = CharacterCasing.Normal
                ObjPO.IsiDGVIPOBySupplierName(Me.dgvItemIPO, Me.TxtIPONo.Text)
                DTA.FillBy_PrimaryPONsupllierName(TBL, vPrimaryPO, Me.TxtIPONo.Text + "%")
            End If

            Dim rowIPO As DataGridViewRow
            Dim rowPO As DataGridViewRow
            Dim Ketemu As Boolean = False
            If Me.dgvItem.RowCount > 0 Then
                'Lakukan penyocokan dengan DGVItem
                For Each rowPO In Me.dgvItem.Rows
                    Ketemu = False
                    For Each row In TBL.Rows
                        If row.Primary_IPO = Me.dgvItem.Item(Me.ColPrimaryIPO.Index, rowPO.Index).Value And _
                        row.Primary_Product = Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, rowPO.Index).Value Then
                            row.Qty -= Me.dgvItem.Item(Me.ColQty.Index, rowPO.Index).Value
                            Ketemu = True
                            Exit For
                        End If
                    Next
                    If Ketemu = False Then
                        Dim DelDate As Date = #1/1/1950#
                        If CStr(rowPO.Cells(Me.ColDeliveryDate.Index).Value) <> "-" Then
                            DelDate = CDate(rowPO.Cells(Me.ColDeliveryDate.Index).Value)
                        End If
                        TBL.Rows.Add(TBL.Rows.Count, CDec(Me.TxtPrimaryPurchaseOrder.Text), rowPO.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        rowPO.Cells(Me.ColPrimaryIPO.Index).Value, rowPO.Cells(Me.ColQty.Index).Value, rowPO.Cells(Me.ColPrimaryUnits.Index).Value, _
                        rowPO.Cells(Me.ColDeliveryCont.Index).Value, DelDate, _
                        rowPO.Cells(Me.ColUnitPrice.Index).Value, rowPO.Cells(Me.ColDiscount.Index).Value, rowPO.Cells(Me.ColNettPrice.Index).Value, _
                        rowPO.Cells(Me.ColAmount.Index).Value, rowPO.Cells(Me.ColRemark.Index).Value, _
                        rowPO.Cells(Me.colUserSecondaryUnits.Index).Value, rowPO.Cells(Me.ColCodeNo.Index).Value, rowPO.Cells(Me.ColDescription.Index).Value, _
                        rowPO.Cells(Me.ColUnit.Index).Value, rowPO.Cells(Me.ColIPONo.Index).Value, 0, "", "")
                    End If
                Next

                For i As Integer = 0 To TBL.Rows.Count - 1
                    If i = TBL.Rows.Count Then
                        Exit For
                    End If
                    If TBL.Rows(i).Item(TBL.Columns.IndexOf(TBL.QtyColumn)) <= 0 Then
                        TBL.Rows.RemoveAt(i)
                        i = -1
                    End If
                Next
                'Lakukan penyocokan dengan IPO
                For Each row In TBL.Rows
                    Ketemu = False
                    For Each rowIPO In Me.dgvItemIPO.Rows
                        If row.Primary_IPO = Me.dgvItemIPO.Item(Me.colPrimaryIPODGVIPO.Index, rowIPO.Index).Value And _
                        row.Primary_Product = Me.dgvItemIPO.Item(Me.colPrimaryCodeNoDGVIPO.Index, rowIPO.Index).Value Then
                            Me.dgvItemIPO.Item(Me.colQtyDGVIPO.Index, rowIPO.Index).Value -= row.Qty
                            Ketemu = True
                            Exit For
                        End If
                    Next
                    If Ketemu = False Then
                        For Each rowPO In Me.dgvItem.Rows
                            If row.Primary_IPO = Me.dgvItem.Item(Me.ColPrimaryIPO.Index, rowPO.Index).Value And _
                               row.Primary_Product = Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, rowPO.Index).Value Then
                                Ketemu = True
                            End If
                        Next
                        If Ketemu = False Then
                            Me.dgvItemIPO.Rows.Add("0", row.No_IPO, row.Product_Code, row.Description, row.Qty, row.Unit_Code, _
                            row.Unit_Price, row.Discount, row.Amount, row.Remarks, row.Primary_Product, row.Primary_Units, row.Primary_IPO, row.Use_Secondary, row.Delivery_Date)
                        End If
                    End If
                Next
            End If

            For i As Integer = 0 To Me.dgvItemIPO.RowCount
                If i = Me.dgvItemIPO.RowCount Then
                    Exit For
                End If
                If Me.dgvItemIPO.Item(Me.colQtyDGVIPO.Index, i).Value <= 0 Then
                    Me.dgvItemIPO.Rows.RemoveAt(i)
                    i -= 1
                End If
            Next

            'set number
            ObjCD.SetNumber(Me.dgvItemIPO)
            'digit grouping
            For Each rowIPO In Me.dgvItemIPO.Rows
                rowIPO.Cells(Me.colQtyDGVIPO.Index).Value = Me.ObjCD.DigitGrouping(rowIPO.Cells(Me.colQtyDGVIPO.Index).Value)
                rowIPO.Cells(Me.colUnitPriceDGVIPO.Index).Value = Me.ObjCD.DigitGrouping(rowIPO.Cells(Me.colUnitPriceDGVIPO.Index).Value)
                rowIPO.Cells(Me.colAmountDGVIPO.Index).Value = Me.ObjCD.DigitGrouping(rowIPO.Cells(Me.colAmountDGVIPO.Index).Value)
                If CDate(rowIPO.Cells(Me.ColDeliveryTimeIPO.Index).Value) = #1/1/1950# Then
                    rowIPO.Cells(Me.ColDeliveryTimeIPO.Index).Value = "-"
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub SamakanHargaDiskonNettPriceUnitPriceDGVIPOdanDGVItem(ByVal rowPeriksa As DataGridViewRow, _
    ByVal KirimDariIPO As Boolean)
        Try
            Dim row As DataGridViewRow
            If KirimDariIPO = False Then
                'DGV Item 
                For Each row In Me.dgvItem.Rows
                    If row.Index <> rowPeriksa.Index Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = rowPeriksa.Cells(Me.ColPrimaryCodeNo.Index).Value Then
                            If row.Cells(Me.ColDiscount.Index).Value <> rowPeriksa.Cells(Me.ColDiscount.Index).Value Or _
                               row.Cells(Me.ColNettPrice.Index).Value <> rowPeriksa.Cells(Me.ColNettPrice.Index).Value Or _
                               row.Cells(Me.ColUnitPrice.Index).Value <> rowPeriksa.Cells(Me.ColUnitPrice.Index).Value Then
                                row.Cells(Me.ColDiscount.Index).Value = rowPeriksa.Cells(Me.ColDiscount.Index).Value
                                row.Cells(Me.ColNettPrice.Index).Value = rowPeriksa.Cells(Me.ColNettPrice.Index).Value
                                row.Cells(Me.ColUnitPrice.Index).Value = rowPeriksa.Cells(Me.ColUnitPrice.Index).Value
                                Me.Hitung_Amount(row)
                            End If
                        End If
                    End If
                Next
            Else
                'DGVIPO cek dengan DGV Item
                For Each row In Me.dgvItem.Rows
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = rowPeriksa.Cells(Me.colPrimaryCodeNoDGVIPO.Index).Value Then
                        If row.Cells(Me.ColDiscount.Index).Value <> rowPeriksa.Cells(Me.colDiscountDGVIPO.Index).Value Or _
                           row.Cells(Me.ColUnitPrice.Index).Value <> rowPeriksa.Cells(Me.colUnitPriceDGVIPO.Index).Value Then
                            rowPeriksa.Cells(Me.colDiscountDGVIPO.Index).Value = row.Cells(Me.ColDiscount.Index).Value
                            rowPeriksa.Cells(Me.colUnitPriceDGVIPO.Index).Value = row.Cells(Me.ColUnitPrice.Index).Value
                            '    Me.Hitung_Amount(rowPeriksa)
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Dim Index As Integer = Me.dgvItem.SelectedRows(0).Index
        Try

            ''@rosadinnn
            ''@MUaT
            If Me.ObjPO.ApakahBolehHapusItem(Me.TxtPrimaryPurchaseOrder.Text, Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.Row.Index).Value) = False Then
                e.Cancel = True
                Exit Sub
            End If

            Me.SamakanHargaDiskonNettPriceUnitPriceDGVIPOdanDGVItem(Me.dgvItem.SelectedRows(0), False)
            If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColPrimaryIPO.Index, Index).Value) = False Then
                'Cek hanya yang ada IPO
                If CDec(Me.dgvItem.Item(Me.ColPrimaryIPO.Index, Index).Value) >= 0 Then
                    Dim Ketemu As Boolean = False
                    '                For Each row In Me.dgvItemIPO.Rows
                    ' If row.Cells(Me.colPrimaryCodeNoDGVIPO.Index).Value = Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, Index).Value And _
                    ' row.Cells(Me.colPrimaryIPODGVIPO.Index).Value = Me.dgvItem.Item(Me.ColPrimaryIPO.Index, Index).Value Then
                    ' row.Cells(Me.colQtyDGVIPO.Index).Value = CDec(row.Cells(Me.colQtyDGVIPO.Index).Value) + CDec(Me.dgvItem.Item(Me.ColQty.Index, Index).Value)
                    ' Ketemu = True
                    ' Exit For
                    ' End If
                    '     Next
                    If Ketemu = False Then
                        Me.IPONo = e.Row.Cells(Me.ColIPONo.Index).Value
                    Else
                        Me.IPONo = ""
                    End If
                End If
            End If

            '4. hapuskan item delivery
            Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + Me.dgvItem.Item(Me.ColIdentifier.Index, e.Row.Index).Value.ToString + "'"
            While Me.vTBLItemDelivery.DefaultView.Count > 0
                Me.vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(0).Row
                Me.vTBLItemDelivery.Rows.Remove(Me.vrowItemDelivery)
            End While
            Me.vTBLItemDelivery.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnKanan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKanan.Click
        If Me.dgvItemIPO.SelectedRows.Count > 0 Then
            Me.SamakanHargaDiskonNettPriceUnitPriceDGVIPOdanDGVItem(Me.dgvItemIPO.SelectedRows(0), True)
            Me.PindahkanItemIPOKePO(Me.dgvItemIPO.SelectedRows(0))
            Me.dgvItemIPO.Rows.RemoveAt(Me.dgvItemIPO.SelectedRows(0).Index)
            Me.Hitung_AmountSemuaBaris()
            Me.Hitung_Total()
        End If
    End Sub

    Private Sub Hitung_AmountSemuaBaris()
        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Dim row As DataGridViewRow
            'hitung nett price
            For Each row In Me.dgvItem.Rows
                row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColUnitPrice.Index).Value * _
                      (100 - row.Cells(Me.ColDiscount.Index).Value) / 100, _
                      DefaultPrecision))
                'hitung amount
                If row.Cells(Me.colUserSecondaryUnits.Index).Value = False Then
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColQty.Index).Value * _
                    row.Cells(Me.ColNettPrice.Index).Value, DefaultPrecision))
                Else
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(0, DefaultPrecision))
                End If
            Next

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PindahkanItemIPOKePO(ByVal rowIPO As DataGridViewRow)
        Dim rowItem As DataGridViewRow
        Dim Ketemu As Boolean = False
        Try
            If Ketemu = False Then
                Dim DelDateCont As Boolean = False
                Dim DelDate As String = "-"
                For Each rowItem In Me.dgvItem.Rows
                    If rowItem.Cells(Me.ColCodeNo.Index).Value = rowIPO.Cells(Me.colCodeNoDGVIPO.Index).Value Then
                        DelDateCont = rowItem.Cells(Me.ColDeliveryCont.Index).Value
                        DelDate = rowItem.Cells(Me.ColDeliveryDate.Index).Value
                        Exit For
                    End If
                Next
                Me.dgvItem.Rows.Add("0", rowIPO.Cells(Me.colCodeNoDGVIPO.Index).Value, _
                rowIPO.Cells(Me.colDescriptionDGVIPO.Index).Value, _
                rowIPO.Cells(Me.colIPONoDGVIPO.Index).Value, Me.ObjCD.DigitGrouping(rowIPO.Cells(Me.colQtyDGVIPO.Index).Value), _
                rowIPO.Cells(Me.colUnitDGVIPO.Index).Value, Me.ObjCD.DigitGrouping(rowIPO.Cells(Me.colUnitPriceDGVIPO.Index).Value), _
                rowIPO.Cells(Me.colDiscountDGVIPO.Index).Value, Me.ObjCD.DigitGrouping(rowIPO.Cells(Me.ColNettPrice.Index).Value), _
                Me.ObjCD.DigitGrouping(rowIPO.Cells(Me.colAmountDGVIPO.Index).Value), DelDateCont, _
                DelDate, rowIPO.Cells(Me.colRemarksDGVIPO.Index).Value, rowIPO.Cells(Me.colPrimaryCodeNoDGVIPO.Index).Value, _
                rowIPO.Cells(Me.colPrimaryUnitsDGVIPO.Index).Value, rowIPO.Cells(Me.colPrimaryIPODGVIPO.Index).Value, _
                rowIPO.Cells(Me.colUseSecondaryUnitsDGVIPO.Index).Value)
            End If
            ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnKananSemua_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKananSemua.Click
        Try
            If Me.dgvItemIPO.SelectedRows.Count > 0 Then
                Dim row As DataGridViewRow
                For Each row In Me.dgvItemIPO.Rows
                    Me.SamakanHargaDiskonNettPriceUnitPriceDGVIPOdanDGVItem(row, True)
                    Me.PindahkanItemIPOKePO(row)
                Next
                Me.dgvItemIPO.Rows.Clear()
                Me.Hitung_AmountSemuaBaris()
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindIPO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindIPO.Click
        Try
            If Me.cboFilterIPO.SelectedItem = "Code No" Then
                Dim i As New Core_Module.frmCariMasterProduct
                i.BackColorAll = Me.BAckColorAll
                i.txtCari.Text = Me.TxtIPONo.Text
                i.cboCriteria.SelectedIndex = 0
                i.ShowDialog()
                If i.PrimaryProduct > -1 Then
                    Me.TxtIPONo.Text = ObjProduct.GetProductCode(i.PrimaryProduct)
                End If
                i.Dispose()
            ElseIf Me.cboFilterIPO.SelectedItem = "Description" Then
                Dim i As New Core_Module.frmCariMasterProduct
                i.BackColorAll = Me.BAckColorAll
                i.txtCari.Text = Me.TxtIPONo.Text
                i.cboCriteria.SelectedIndex = 1
                i.ShowDialog()
                If i.PrimaryProduct > -1 Then
                    Me.TxtIPONo.Text = ObjProduct.GetDescription(i.PrimaryProduct)
                End If
                i.Dispose()
            ElseIf Me.cboFilterIPO.SelectedItem = "IPO No" Then

            ElseIf Me.cboFilterIPO.SelectedItem = "No Filter" Then

            ElseIf Me.cboFilterIPO.SelectedItem = "Supplier Code" Then
                Dim i As New Core_Module.FrmCariMasterSupplier
                i.BackColorAll = Me.BAckColorAll
                i.txtCari.Text = Me.TxtIPONo.Text
                i.cboCriteria.SelectedIndex = 1
                i.ShowDialog()
                If i.PrimarySupplier > -1 Then
                    Me.TxtIPONo.Text = ObjSupplier.GetSupplierCode(i.PrimarySupplier)
                End If
                i.Dispose()
            ElseIf Me.cboFilterIPO.SelectedItem = "Supplier Name" Then
                Dim i As New Core_Module.FrmCariMasterSupplier
                i.BackColorAll = Me.BAckColorAll
                i.txtCari.Text = Me.TxtIPONo.Text
                i.cboCriteria.SelectedIndex = 1
                i.ShowDialog()
                If i.PrimarySupplier > -1 Then
                    Me.TxtIPONo.Text = ObjSupplier.GetSupplierName(i.PrimarySupplier)
                End If
                i.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtIPONo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtIPONo.TextChanged
        Me.LoadDGVIPODanPO(Me.TxtPrimaryPurchaseOrder.Text)
    End Sub

    Private Sub cboFilterIPO_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFilterIPO.SelectedIndexChanged
        If Me.cboFilterIPO.SelectedIndex > -1 Then
            Me.LoadDGVIPODanPO(Me.TxtPrimaryPurchaseOrder.Text)
        End If
    End Sub

    Private Sub TxtIPONo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtIPONo.KeyDown
        Try
            If Me.cboFilterIPO.SelectedIndex > -1 Then
                If e.KeyCode = Keys.Enter Then
                    Me.LoadDGVIPODanPO(Me.TxtPrimaryPurchaseOrder.Text)
                ElseIf e.KeyValue = 40 And Me.dgvItemIPO.RowCount > 0 Then
                    Me.dgvItemIPO.Focus()
                    If Me.dgvItemIPO.Rows.Count > 0 Then
                        Me.dgvItemIPO.Rows.Item(0).Selected = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbHistoricalPrice_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbHistoricalPrice.LinkClicked
        Dim ij As New FrmTransaksiPO_HistoricalPrice

        Try
            If Me.TxtPrimaryCodeNo.Text <> "" And Me.TxtPrimaryCodeNo.Text <> "-1" Then
                Me.ObjPO.ExplainHistoricalPrice(Me.TxtPrimaryCodeNo.Text, Me.TBL, ij.dgvPrice, Me.ObjCD)
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbDetailSign.LinkClicked
        If Me.TxtRemark.Enabled = False Then
            If Me.TxtPrimaryPurchaseOrder.Text = "" Or Me.TxtPrimaryPurchaseOrder.Text = "-1" Then
                Exit Sub
            End If
        End If

        Dim ij As New FrmTransaksiPO_DetailSIgn
        ij.BackColorAll = Me.BAckColorAll

        Try
            If Me.TxtRemark.Enabled = False Then
                ij.TxtPreparedBy.ReadOnly = True
                ij.TxtPreparedBy_Jabatan.ReadOnly = True
                ij.TxtApprovedBy.ReadOnly = True
                ij.TxtApprovedBy_Jabatan.ReadOnly = True
                ij.TxtAcknowledgeBy.ReadOnly = True
                ij.TxtAcknowledgeBy_Jabatan.ReadOnly = True
                ij.btnOK.Visible = False
            End If

            ij.TxtPreparedBy.Text = Me.ObjPO.PreparedBy
            ij.TxtPreparedBy_Jabatan.Text = Me.ObjPO.PreparedBy_Jabatan
            ij.TxtApprovedBy.Text = Me.ObjPO.ApprovedBy
            ij.TxtApprovedBy_Jabatan.Text = Me.ObjPO.ApprovedBy_Jabatan
            ij.TxtAcknowledgeBy.Text = Me.ObjPO.AcknowledgeBy
            ij.TxtAcknowledgeBy_Jabatan.Text = Me.ObjPO.AcknowledgeBy_Jabatan

            ij.ShowDialog()

            If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                Me.ObjPO.PreparedBy = ij.TxtPreparedBy.Text
                Me.ObjPO.PreparedBy_Jabatan = ij.TxtPreparedBy_Jabatan.Text
                Me.ObjPO.ApprovedBy = ij.TxtApprovedBy.Text
                Me.ObjPO.ApprovedBy_Jabatan = ij.TxtApprovedBy_Jabatan.Text
                Me.ObjPO.AcknowledgeBy = ij.TxtAcknowledgeBy.Text
                Me.ObjPO.AcknowledgeBy_Jabatan = ij.TxtAcknowledgeBy_Jabatan.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Sub

    Private Sub dtpDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.Enter
        Try
            Me.Buffer = Format(Me.dtpDate.Value, "MMM yyyy")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.Leave
        Try
            If Me.TxtPrimaryCodeNo.Text <> "" And Me.TxtPrimaryCodeNo.Text <> "-1" Then

                ''@rosadinnn
                ''MUaT
                Dim DefPrec_ As Integer = 0
                If Me.ObjPO.IsMultiUnitProduct(Me.TxtPrimaryCodeNo.Text) = False Then
                    DefPrec_ = Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)
                Else
                    DefPrec_ = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text)
                End If

                Me.TxtQtyperBatch.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatch.Text, DefPrec_)
                Me.Hitung_Amount()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtUnitPrice, sender, e, Me.DefaultPrecision, Me.ObjCD)
            'Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(Me.TxtUnitPrice.Text, Me.DefaultPrecision)
            'unit price harus selalu lebih besar dari disc amount
            'hanya berlaku kalo ngitung disc berdasarkan disc amount, klo berdasarkan perc, langsung hitung ulang disc amountnya
            If Me.TxtDiscFromPerct.Text = "0" Then
                If CDec(Me.TxtDiscAmount.Text) > CDec(Me.TxtUnitPrice.Text) Then
                    Me.TxtUnitPrice.Text = Me.Buffer
                    Exit Sub
                End If
            End If
            Me.Hitung_Amount()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbShipTo.LinkClicked
        Dim ij As New FrmTransaksiPO_DetailShipTo
        Dim DV As Data.DataRowView
        Dim row As Core_Module.DseMasterCompanyData.tbltrans_company_data_itemRow
        Dim i As Integer = 0

        Try
            If Me.TxtPONo.Text = "" Then
                Exit Sub
            End If

            ij.vAddIdentifier = Me.TxtIdentifierShipTo.Text
            ij.Text = "Ship To Details"
            If Me.TxtRemark.Enabled = True Then
                ij.btnOK.BringToFront()
            Else
                ij.cboCustAddCode.Enabled = False
            End If
            ij.btnExit.TabStop = False
            ij.ShowDialog()

            If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                Me.ObjCompanyData.IsiCboAdd(Me.TBL_Add, Me.cboShipTo)
                For Each DV In Me.TBL_Add.DefaultView
                    row = DV.Row

                    If row.Add_Identifier = ij.vAddIdentifierResult Then
                        Me.cboShipTo.SelectedIndex = i
                        Exit For
                    End If

                    i += 1
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtDiscount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtDiscount, sender, e, 2, Me.ObjCD)
            If CDec(Me.TxtDiscount.Text) <> CDec(Me.Buffer) Then
                Me.TxtDiscFromPerct.Text = "1"
                Me.Hitung_Amount()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboShipTo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboShipTo.SelectedValueChanged
        Dim row As Core_Module.DseMasterCompanyData.tbltrans_company_data_itemRow

        Try
            If Me.cboShipTo.SelectedIndex = -1 Then
                Me.TxtIdentifierShipTo.Text = ""
            Else
                row = Me.TBL_Add.DefaultView(Me.cboShipTo.SelectedIndex).Row
                Me.TxtIdentifierShipTo.Text = row.Add_Identifier
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub TxtPONo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPONo.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPONo.ReadOnly = True Then
                    Me.TxtPONo.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtPONo.BackColor = Me.BAckColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPONo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPONo.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPONo.ReadOnly = True Then
                    Me.TxtPONo.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtPONo.BackColor = Me.TxtPrimaryPurchaseOrder.BackColor
                End If
            End If
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
        Try
            If Me.TxtRemark.Enabled = True And Me.txtPrimarySupplier.Text <> "" And Me.txtPrimarySupplier.Text <> "-1" Then
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = False Then
                    Exit Sub
                End If
                If Me.baru = False Then
                    If Me.ObjPO.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.ObjPO.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtPONo.Text = Me.ObjPO.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtPONo.Text = Me.ObjPO.GetNoBaru1(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                    End If
                Else
                    Me.TxtPONo.Text = Me.ObjPO.GetNoBaru1(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                End If

                Me.dtpDeliveryDate.Value = Me.dtpDate.Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryMax(ByVal vTBL As DseTransaksiPO.tbltrans_po_itemdeliveryDataTable) As Double
        Dim Retval As Double = 0

        Try
            For Each Me.vrowItemDelivery In Me.vTBLItemDelivery.Rows
                If Retval < Me.vrowItemDelivery.primary Then
                    Retval = Me.vrowItemDelivery.primary
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtDiscAmount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.Enter
        Try
            Me.Buffer = Me.TxtDiscAmount.Text
            Me.ObjCD.Number_Enter(Me.TxtDiscAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtDiscAmount.ReadOnly = False Then
                    Me.TxtDiscAmount.BackColor = Me.BAckColorAll
                Else
                    Me.TxtDiscAmount.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscAmount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDiscAmount, sender, e, CDec(Me.TxtUnitPrice.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscAmount.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDiscAmount, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtDiscAmount, sender, e, Me.DefaultPrecision, Me.ObjCD)
            'disc amount tidak boleh lebih besar dari unit price
            If CDec(Me.TxtDiscAmount.Text) > CDec(Me.TxtUnitPrice.Text) Then
                Me.TxtDiscAmount.Text = Me.Buffer
                Exit Sub
            End If
            If CDec(Me.TxtDiscAmount.Text) <> CDec(Me.Buffer) Then
                Me.TxtDiscFromPerct.Text = "0"
                Me.Hitung_Amount()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtDiscAmount.ReadOnly = False Then
                    Me.TxtDiscAmount.BackColor = Me.TxtPrimaryPurchaseOrder.BackColor
                Else
                    Me.TxtDiscAmount.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub PrepareMultiUnit(ByVal vPrimaryProduct As Double)
        Dim dtaMultiUnit As New DseTransaksiPOTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Try
            Me.CboMultiUnit.DataSource = Nothing
            If Me.ObjPO.IsMultiUnitProduct(vPrimaryProduct) = False Then
                Me.lblUnits.Visible = True
                Me.CboMultiUnit.Visible = False
                Me.CboMultiUnit.TabStop = False
                Me.LblUnitPrice.Text = "Unit Price/" & lblUnits.Text
            Else
                Me.lblUnits.Visible = False
                Me.CboMultiUnit.Visible = True
                Me.CboMultiUnit.TabStop = True
                dtaMultiUnit.Connection.ConnectionString = Core_Module.GetConstring
                dtaMultiUnit.FillBy_PrimaryProduct(Me.tblMultiUnit, Me.TxtPrimaryCodeNo.Text)
                Me.ObjPO.IsiCboUnit(vPrimaryProduct, Me.CboMultiUnit, Me.tblCboUnit, Me.tblMultiUnit, Me.txtPrimaryUnits)
                Me.CboMultiUnit.SelectedIndex = -1
                Me.CboMultiUnit.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboMultiUnit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboMultiUnit.Leave
        Try
            If Me.TxtPrimaryCodeNo.Text <> "" And Me.TxtPrimaryCodeNo.Text <> "-1" Then

                ''@rosadinnn
                ''MUaT
                Dim DefPrec_ As Integer = 0
                If Me.ObjPO.IsMultiUnitProduct(Me.TxtPrimaryCodeNo.Text) = False Then
                    DefPrec_ = Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)
                Else
                    DefPrec_ = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text)
                End If

                Me.TxtQtyperBatch.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatch.Text, DefPrec_)
                Me.Hitung_Amount()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboMultiUnit_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboMultiUnit.SelectedValueChanged
        Dim UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc As Decimal
        Dim row As DseTransaksiPO.CboUnitRow

        Try
            If Me.CboMultiUnit.SelectedIndex = -1 Then
                If Me.ObjPO.IsMultiUnitProduct(Me.TxtPrimaryCodeNo.Text) = True Then
                    Me.txtPrimaryUnits.Text = "-1"
                End If
            Else
                row = Me.tblCboUnit.DefaultView(Me.CboMultiUnit.SelectedIndex).Row
                Me.txtPrimaryUnits.Text = row.Primary
                Me.LblUnitPrice.Text = "Unit Price/" & Me.CboMultiUnit.Text
                Me.lblUnits.Text = Me.CboMultiUnit.Text

                Me.TxtQtyperBatch.Text = ObjCD.DigitGrouping(Me.TxtQtyperBatch.Text, ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text))

                Me.ObjPO.GetHistoricalPrice(Me.txtPrimaryUnits.Text, Me.TxtPrimaryCodeNo.Text, _
                Me.txtPrimaryCurrency.Text, UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc)

                Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.DefaultPrecision)
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                Me.TxtNettPrice.Text = Me.ObjCD.DigitGrouping(NettPrice, Me.DefaultPrecision)
                Me.TxtDiscAmount.Text = Me.ObjCD.DigitGrouping(DiscAmount, Me.DefaultPrecision)
                Me.TxtDiscFromPerct.Text = DiscFromPerc

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetColQtyDefault(ByVal vRow As DataGridViewRow)
        Try
            vRow.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(CDec(vRow.Cells(Me.ColQtyMulti.Index).Value) * _
                                                CDec(vRow.Cells(Me.ColRatioDefault.Index).Value) / _
                                                CDec(vRow.Cells(Me.ColRatioReference.Index).Value), _
                                                Me.ObjProduct.GetAngkaDecimal(vRow.Cells(Me.ColPrimaryCodeNo.Index).Value))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetQtyDefault(ByVal vQtyMulti As Decimal, ByVal vRatioDefault As Decimal, ByVal vRatioReference As Decimal) As Decimal
        Try
            Return CDec(vQtyMulti * vRatioDefault / vRatioReference)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Sub SetRatioAndQtyDefault(ByVal vDgvRow As DataGridViewRow, ByVal vSetQtyDefault As Boolean)
        Try
            With vDgvRow
                If .Cells(Me.ColPrimaryUnitMulti.Index).Value <> .Cells(Me.ColPrimaryUnits.Index).Value Then
                    .Cells(Me.ColRatioDesc.Index).Value = .Cells(Me.ColRatioReference.Index).Value & " " & .Cells(Me.ColUnitCodeMulti.Index).Value & " = " & _
                    .Cells(Me.ColRatioDefault.Index).Value & " " & .Cells(Me.ColUnit.Index).Value
                Else
                    .Cells(Me.ColRatioDesc.Index).Value = "-"
                End If
                
                If vSetQtyDefault = True Then
                    Me.SetColQtyDefault(vDgvRow)
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbColumnSettings_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbColumnSettings.LinkClicked
        Dim ij As New FrmTransaksiPO_SetVisible

        Try
            'tampilkan form untuk setting data
            ij.vUserName = Me.UserCode
            ij.IsMultiUnit = Me.IsMultiUnit
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                'update kolom
                Me.UpdateKolomPO()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub UpdateKolomPO()
        Dim Dummy As String = ""
        Dim IsMultiUnit_ As Boolean = Me.IsMultiUnit

        Try
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowQtyDefault, Me.UserCode)
            If Dummy = "" Then
                Me.ColQty.Visible = False
            Else
                If Dummy = 1 Then
                    Me.ColQty.Visible = True
                    If IsMultiUnit_ = False Then
                        Me.ColQty.Visible = False
                    End If
                Else
                    Me.ColQty.Visible = False
                End If
            End If

            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowUnitCodeDefault, Me.UserCode)
            If Dummy = "" Then
                Me.ColUnit.Visible = False
            Else
                If Dummy = 1 Then
                    Me.ColUnit.Visible = True
                    If IsMultiUnit_ = False Then
                        Me.ColUnit.Visible = False
                    End If
                Else
                    Me.ColUnit.Visible = False
                End If
            End If

            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowToleranceDown, Me.UserCode)
            If Dummy = "" Then
                Me.ColToleranceDown.Visible = False
                Me.ColToleranceDownPerc.Visible = False
            Else
                If Dummy = 1 Then
                    Me.ColToleranceDown.Visible = True
                    Me.ColToleranceDownPerc.Visible = True
                Else
                    Me.ColToleranceDown.Visible = False
                    Me.ColToleranceDownPerc.Visible = False
                End If
            End If

            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
                                    Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowToleranceUp, Me.UserCode)
            If Dummy = "" Then
                Me.ColToleranceUp.Visible = False
                Me.ColToleranceUpPerc.Visible = False
            Else
                If Dummy = 1 Then
                    Me.ColToleranceUp.Visible = True
                    Me.ColToleranceUpPerc.Visible = True
                Else
                    Me.ColToleranceUp.Visible = False
                    Me.ColToleranceUpPerc.Visible = False
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetTolerancePerc(ByVal vRowIndex As Integer)
        Dim RowDgv As DataGridViewRow
        Try
            RowDgv = Me.dgvItem.Rows(vRowIndex)
            With RowDgv
                If CDec(.Cells(Me.ColQtyMulti.Index).Value) > 0 Then
                    .Cells(Me.ColToleranceDownPerc.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColToleranceDown.Index).Value / .Cells(Me.ColQtyMulti.Index).Value * 100, 2)
                Else
                    .Cells(Me.ColToleranceDownPerc.Index).Value = ObjCD.DigitGrouping(0, 2)
                End If
                If CDec(.Cells(Me.ColQtyMulti.Index).Value) > 0 Then
                    .Cells(Me.ColToleranceUpPerc.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColToleranceUp.Index).Value / .Cells(Me.ColQtyMulti.Index).Value * 100, 2)
                Else
                    .Cells(Me.ColToleranceUpPerc.Index).Value = ObjCD.DigitGrouping(0, 2)
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetToleranceVal(ByVal vRowIndex As Integer)
        Dim RowDgv As DataGridViewRow
        Try
            RowDgv = Me.dgvItem.Rows(vRowIndex)
            With RowDgv
                .Cells(Me.ColToleranceDown.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColToleranceDownPerc.Index).Value / 100 * .Cells(Me.ColQtyMulti.Index).Value, ObjProduct.GetAngkaDecimal(.Cells(ColPrimaryCodeNo.Index).Value, .Cells(ColPrimaryUnitMulti.Index).Value))
                .Cells(Me.ColToleranceUp.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColToleranceUpPerc.Index).Value / 100 * .Cells(Me.ColQtyMulti.Index).Value, ObjProduct.GetAngkaDecimal(.Cells(ColPrimaryCodeNo.Index).Value, .Cells(ColPrimaryUnitMulti.Index).Value))
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDgv()
        Dim isMultiUnit As Boolean = False
        Dim Dummy As String = ""
        Try
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                With row
                    If .Cells(Me.ColPrimaryUnits.Index).Value <> .Cells(Me.ColPrimaryUnitMulti.Index).Value Then
                        isMultiUnit = True
                    End If
                End With
            Next
            If isMultiUnit = False Then
                Me.ColQty.Visible = False
                Me.ColUnit.Visible = False
                Me.ColRatioDesc.Visible = False
            Else
                Me.ColQty.Visible = True
                Me.ColUnit.Visible = True

                ''walaupun multiunit=true tapi jika visible=false user defined jgn ditampilkan
                Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
                            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowQtyDefault, Me.UserCode)
                If Dummy = "" Then
                    Me.ColQty.Visible = False
                Else
                    If Dummy <> 1 Then
                        Me.ColQty.Visible = False
                    End If
                End If
                Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.PurchaseOrder, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.PurchaseOrder_ShowUnitCodeDefault, Me.UserCode)
                If Dummy = "" Then
                    Me.ColUnit.Visible = False
                Else
                    If Dummy <> 1 Then
                        Me.ColUnit.Visible = False
                    End If
                End If

                Me.ColRatioDesc.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsMultiUnit() As Boolean
        Try
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                With row
                    If .Cells(Me.ColPrimaryUnits.Index).Value <> .Cells(Me.ColPrimaryUnitMulti.Index).Value Then
                        Return True
                    End If
                End With
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub NewPO()
        Dim Dummy As String = ""
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'pastikan cara penomoran transaksi untuk po lokal dan po impor sama
            If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor) <> _
            Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) Then
                MsgBox("Please specify similar numbering convention for local and import purchase order", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.ObjPO.ResetObject(Me.dgvItem)
            'secara default yang melakukan tanda tangan pada PO adalah, ambil dari setting PO
            Me.ObjPO.PreparedBy = My.Settings.PO_Prepared_By
            Me.ObjPO.PreparedBy_Jabatan = My.Settings.PO_Prepared_Position
            Me.ObjPO.ApprovedBy = My.Settings.PO_Approved_By
            Me.ObjPO.ApprovedBy_Jabatan = My.Settings.PO_Approved_Position
            Me.ObjPO.AcknowledgeBy = My.Settings.PO_Acknowled_By
            Me.ObjPO.AcknowledgeBy_Jabatan = My.Settings.PO_Acknowledge_Position

            If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                Me.TxtPONo.Text = Me.ObjPO.GetNoBaru1(Me.dtpDate.Value)
            End If
            Me.txtAutohirizationStatus.Text = "Unautorized"

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            'kalo new mata uang langsung IDR
            Me.txtPrimaryCurrency.Text = "1"
            Me.txtCurrencyCode.Text = Me.objCurrency.GetCurrencyCode(1)

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If Dummy = "" Then
                Me.chkTax.Checked = False
                Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
            Else
                Me.chkTax.Checked = True
                Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                If Dummy = "" Then
                    Dummy = "0"
                End If
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(Dummy, 2)
            End If

            'Me.chkTax.Checked = True
            'Me.txtTax.Text = Me.ObjCD.DigitGrouping("10", 2)

            'isikan ulang cbo ship to, dan set index default
            Me.ObjCompanyData.IsiCboAdd(Me.TBL_Add, Me.cboShipTo)
            Me.SetIndexCboShipToDefault()

            'secara default payment term adalah dr invoice date
            Me.cboPaymentDueDateReference.SelectedIndex = 1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub CariPO_ForAdd(ByVal vPrimaryPO As Double)
        Try
            Me.NonAktifkanDetail()
            Me.baru = False
            Me.ObjPO.CariDetail(vPrimaryPO, Me.dgvItem, Me.vTBLItemDelivery, Me)
            Me.SalinKeField()
            Me.AktifkanDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetSupplierData(ByVal vPrimarySupplier As Double)
        Dim SupplierName, CPName, PhoneNo As String
        SupplierName = ""
        CPName = ""
        PhoneNo = ""
        Dim ApakahLokal As Boolean = False
        Dim IndexStr, IndexStr1 As Integer

        Try
            Me.txtPrimarySupplier.Text = vPrimarySupplier
            Me.ObjSupplier.Get_SuppName_ContactPerson_PhoneNo(Me.txtPrimarySupplier.Text, SupplierName, CPName, PhoneNo)
            Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
            Me.TxtSupplierName.Text = SupplierName
            Me.ObjSupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
            Me.cboSupplierAddress.SelectedIndex = 0
            Me.cboSupplierAddress.Focus()
            Me.txtContactPerson.Text = CPName
            Me.txtPhoneNo.Text = PhoneNo

            If Me.baru = True Then
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                    Me.TxtPONo.Text = Me.ObjPO.GetNoBaru1(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                End If
            Else
                ApakahLokal = Me.ObjSupplier.ApakahLokal(Me.txtPrimarySupplier.Text)
                If ApakahLokal = True Then
                    IndexStr = Me.TxtPONo.Text.IndexOf("L")
                    IndexStr1 = Me.TxtPONo.Text.IndexOf("I")
                Else
                    IndexStr = Me.TxtPONo.Text.IndexOf("I")
                    IndexStr1 = Me.TxtPONo.Text.IndexOf("L")
                End If
                If IndexStr = -1 Then
                    If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                        Me.TxtPONo.Text = Me.ObjPO.GetNoBaru1(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                    End If
                End If
            End If

            Me.cboPaymentDueDateReference.SelectedIndex = 1
            Me.txtPaymentTerm.Text = Me.ObjSupplier.GetPaymentTerm(Me.txtPrimarySupplier.Text)
            Me.txtPaymentDueDate.Text = Me.txtPaymentTerm.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetCurrency(ByVal vPrimaryCurrency As Double)
        Try
            Me.txtPrimaryCurrency.Text = vPrimaryCurrency
            Me.txtCurrencyCode.Text = Me.objCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
            Me.SetDefaultPrecision()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStopProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStopProcess.Click
        Try
            If MsgBox("Stop creating po?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.Abort
                Me.LangsungClose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUpdateDeliveryDate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUpdateDeliveryDate.LinkClicked
        Dim f As New FrmTransaksiPO_SetDeliveryDate
        Dim rowD As DseTransaksiPO.tbltrans_po_itemdeliveryRow
        Try
            If Me.dgvItem.RowCount = 0 Or Me.btnOK.Enabled = False Then
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            If f.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.vTBLItemDelivery.Clear()
                If f.RadioButton1.Checked = True Then
                    For Each r As DataGridViewRow In Me.dgvItem.Rows
                        r.Cells(Me.ColDeliveryCont.Index).Value = True
                        r.Cells(Me.ColDeliveryDate.Index).Value = Format(f.DateTimePicker1.Value.Date, "dd MMM yyyy")
                        rowD = Me.vTBLItemDelivery.Newtbltrans_po_itemdeliveryRow
                        With rowD
                            .delivery_date = f.DateTimePicker1.Value.Date
                            .delay = Me.ObjItem.GetPreparationDelay(r.Cells(Me.ColPrimaryCodeNo.Index).Value)
                            .delivery_qty = r.Cells(Me.ColQty.Index).Value
                            .identifier = r.Cells(Me.ColIdentifier.Index).Value
                            .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                            .primary_transaksi = Me.TxtPrimaryPurchaseOrder.Text
                            .remarks = ""
                            .delivery_qty_multi = r.Cells(Me.ColQtyMulti.Index).Value
                        End With
                        Me.vTBLItemDelivery.Rows.Add(rowD)
                    Next
                Else
                    For Each r As DataGridViewRow In Me.dgvItem.Rows
                        r.Cells(Me.ColDeliveryCont.Index).Value = False
                        r.Cells(Me.ColDeliveryDate.Index).Value = "-"
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
End Class