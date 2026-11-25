Option Explicit On
Option Strict Off

Public Class frmLaporanVAT2025_EditCustomer
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public isEdit As Boolean = False
    Public PrimaryCustomer As Double = -1
    Public PrimaryInvoice As Double = -1
    Public TypeTransaction As EnumTransactionType = EnumTransactionType.MasterCustomer
    Public TypeDocument As String = ""
    Public NPWP As String = ""
    Public NoNIKPASS As String = ""
    Dim ObjCustomer As New clsMasterCustomer
    Dim DTA As New DseTrans_SettingNPWPTableAdapters.tblmaster_customerTableAdapter
    Dim TBL As New DseTrans_SettingNPWP.tblmaster_customerDataTable
    Dim DTASales As New DseTrans_SettingNPWPTableAdapters.tbltrans_salesinvoiceTableAdapter
    Dim TBLSales As New DseTrans_SettingNPWP.tbltrans_salesinvoiceDataTable
    Dim DTASalesReturn As New DseTrans_SettingNPWPTableAdapters.tbltrans_salesreturnTableAdapter
    Dim TBLSalesReturn As New DseTrans_SettingNPWP.tbltrans_salesreturnDataTable
    Dim NPWPType As Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP = Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP.Company
    Dim Pilih_NIKPASS As Core_Module.clsMasterCustomer.EnumOptionNikPass = clsMasterCustomer.EnumOptionNikPass.ByNIK

    Public Enum DocumentType As Integer
        NIK
        Passport
        OtherID
    End Enum

    Public Enum EnumTransactionType As Integer
        MasterCustomer
        MasterSupplier
        MasterCompany
        MasterEmployee
        TransactionSales
        TransactionSalesReturn
        TransactionPurchase
        TransactionPurchaseReturn
    End Enum

    Private Sub frmLaporanVAT2025_EditCustomer_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim NPWPSave As String
        Try
            If Me.isEdit = False Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar = True Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        NPWPSave = Me.MTxtNPWP.Text.Trim.Replace(" ", "")
                        'simpan ke master customer
                        DTA.Connection.ConnectionString = Core_Module.GetConstring
                        DTA.UpdateQuery_CustomerTaxData(Me.NPWPType, NPWPSave, Me.TxtAddressPajak1.Text, Me.TxtAddressPajak2.Text, Me.UserCode, _
                        Now, Now, Me.txtNIKPASS.Text, Me.Pilih_NIKPASS, Me.TxtNPWP_Name.Text, Me.TxtTKUID.Text, Me.TxtEmail.Text, Me.PrimaryCustomer)

                        If Me.TypeTransaction = EnumTransactionType.TransactionSales Then
                            'simpan ke transaksi sales invoice
                            DTASales.Connection.ConnectionString = Core_Module.GetConstring
                            DTASales.UpdateQuery_UpdateTaxData(NPWPSave, Me.TxtAddressPajak1.Text, Me.TxtAddressPajak2.Text, Me.Pilih_NIKPASS, Me.txtNIKPASS.Text, _
                            Me.UserCode, Now, Me.PrimaryInvoice)
                        ElseIf Me.TypeTransaction = EnumTransactionType.TransactionSalesReturn Then
                            'simpan ke transaksi sales invoice
                            DTASalesReturn.Connection.ConnectionString = Core_Module.GetConstring
                            DTASalesReturn.UpdateQuery_UpdateTaxData1(Me.UserCode, Now, NPWPSave, Me.TxtAddressPajak1.Text, Me.PrimaryInvoice)
                        End If
                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")

                        If Me.cboTypeNPWP.SelectedIndex = Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP.Company Then
                            Me.TypeDocument = "NPWP"
                        Else
                            Me.TypeDocument = Me.cboDocType.Text
                        End If
                        Me.NPWP = Me.MTxtNPWP.Text
                        Me.NoNIKPASS = Me.txtNIKPASS.Text
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                ElseIf hasil = MsgBoxResult.No Then
                    DialogResult = Windows.Forms.DialogResult.Cancel
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmLaporanVAT2025_EditInvalid_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.txtprimarycustomer.Text = Me.PrimaryCustomer
            If Me.LoadDataCustomer() = False Then
                MsgBox("There is no customer data", MsgBoxStyle.Information, "Information")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function LoadDataCustomer() As Boolean
        Dim row As DseTrans_SettingNPWP.tblmaster_customerRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, Me.PrimaryCustomer)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.txtprimarycustomer.Text = .Primary
                    Me.TxtTaxCode.Text = .Default_Tax_Code
                    Me.MTxtNPWP.Text = .NPWP_No.Replace(" ", "").Replace(".", "").Replace("-", "")
                    Me.TxtNPWP_Name.Text = .NPWP_Name
                    Me.TxtAddressPajak1.Text = .Tax_Add_Line1
                    Me.TxtAddressPajak2.Text = .Tax_Add_Line2
                    Me.TxtTKUID.Text = .TKU_ID
                    Me.TxtEmail.Text = .Email
                    Me.txtNIKPASS.Text = .No_NIKPASS
                    Me.Pilih_NIKPASS = .Pilih_NIKPASS
                    Me.NPWP = .NPWP_No.Replace(" ", "").Replace(".", "").Replace("-", "")
                    Me.NoNIKPASS = .No_NIKPASS
                    Me.cboTypeNPWP.SelectedIndex = .NPWP_Type
                    If .NPWP_Type = Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP.Company Then
                        Me.cboDocType.SelectedIndex = 0
                    Else
                        Me.cboDocType.SelectedIndex = .Pilih_NIKPASS
                    End If
                End With
                Return True
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub cboDocType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDocType.SelectedIndexChanged
        Try
            If Me.cboDocType.SelectedIndex = 0 And Me.txtNIKPASS.Enabled = True Then
                Me.MTxtNPWP.Text = Me.txtNIKPASS.Text
            Else
                Me.txtNIKPASS.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNIKPASS_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNIKPASS.Leave
        Try
            If Me.cboDocType.SelectedIndex = 0 And Me.txtNIKPASS.Enabled = True Then
                Me.MTxtNPWP.Text = Me.txtNIKPASS.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#Region "Txt Focus"
    Private Sub MTxtNPWP_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MTxtNPWP.GotFocus
        Me.MTxtNPWP.BackColor = Me.BackColorAll
    End Sub

    Private Sub MTxtNPWP_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MTxtNPWP.LostFocus
        Me.MTxtNPWP.BackColor = Me.txtprimarycustomer.BackColor
    End Sub

    Private Sub TxtAddressPajak1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddressPajak1.GotFocus
        Me.TxtAddressPajak1.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtAddressPajak1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddressPajak1.LostFocus
        Me.TxtAddressPajak1.BackColor = Me.txtprimarycustomer.BackColor
    End Sub

    Private Sub TxtAddressPajak2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddressPajak2.GotFocus
        Me.TxtAddressPajak2.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtAddressPajak2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddressPajak2.LostFocus
        Me.TxtAddressPajak2.BackColor = Me.txtprimarycustomer.BackColor
    End Sub

    Private Sub TxtEmail_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmail.GotFocus
        Me.TxtEmail.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtEmail_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmail.LostFocus
        Me.TxtEmail.BackColor = Me.txtprimarycustomer.BackColor
    End Sub

    Private Sub txtNIKPASS_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNIKPASS.GotFocus
        Me.txtNIKPASS.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtNIKPASS_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNIKPASS.LostFocus
        Me.txtNIKPASS.BackColor = Me.txtprimarycustomer.BackColor
    End Sub

    Private Sub TxtNPWP_Name_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWP_Name.GotFocus
        Me.TxtNPWP_Name.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtNPWP_Name_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWP_Name.LostFocus
        Me.TxtNPWP_Name.BackColor = Me.txtprimarycustomer.BackColor
    End Sub

    Private Sub TxtTaxCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxCode.GotFocus
        Me.TxtTaxCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtTaxCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxCode.LostFocus
        Me.TxtTaxCode.BackColor = Me.txtprimarycustomer.BackColor
    End Sub

    Private Sub TxtTKUID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTKUID.GotFocus
        Me.TxtTKUID.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtTKUID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTKUID.LostFocus
        Me.TxtTKUID.BackColor = Me.txtprimarycustomer.BackColor
    End Sub
#End Region

    Private Sub BtnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Try
            DialogResult = Windows.Forms.DialogResult.Cancel
            Me.isEdit = False
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim hasil As MsgBoxResult
        Dim NPWPSave As String
        Try
            If Me.ApakahDataBenar = True Then
                hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.cboDocType.SelectedIndex > 0 Then
                        Me.Pilih_NIKPASS = Me.cboDocType.SelectedIndex - 1
                    End If
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    NPWPSave = Me.MTxtNPWP.Text.Trim.Replace(" ", "")

                    'simpan ke master customer
                    DTA.Connection.ConnectionString = Core_Module.GetConstring
                    DTA.UpdateQuery_CustomerTaxData(Me.NPWPType, NPWPSave, Me.TxtAddressPajak1.Text, Me.TxtAddressPajak2.Text, Me.UserCode, _
                    Now, Now, Me.txtNIKPASS.Text, Me.Pilih_NIKPASS, Me.TxtNPWP_Name.Text, Me.TxtTKUID.Text, Me.TxtEmail.Text, Me.PrimaryCustomer)

                    If Me.TypeTransaction = EnumTransactionType.TransactionSales Then
                        'simpan ke transaksi sales invoice
                        DTASales.Connection.ConnectionString = Core_Module.GetConstring
                        DTASales.UpdateQuery_UpdateTaxData(NPWPSave, Me.TxtAddressPajak1.Text, Me.TxtAddressPajak2.Text, Me.Pilih_NIKPASS, Me.txtNIKPASS.Text, _
                        Me.UserCode, Now, Me.PrimaryInvoice)
                    ElseIf Me.TypeTransaction = EnumTransactionType.TransactionSalesReturn Then
                        'simpan ke transaksi sales invoice
                        DTASalesReturn.Connection.ConnectionString = Core_Module.GetConstring
                        DTASalesReturn.UpdateQuery_UpdateTaxData1(Me.UserCode, Now, NPWPSave, Me.TxtAddressPajak1.Text, Me.PrimaryInvoice)
                    End If

                    If Me.cboTypeNPWP.SelectedIndex = Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP.Company Then
                        Me.TypeDocument = "NPWP"
                    Else
                        Me.TypeDocument = Me.cboDocType.Text
                    End If
                    Me.NPWP = Me.MTxtNPWP.Text
                    Me.NoNIKPASS = Me.txtNIKPASS.Text
                    Me.isEdit = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Try
            'Tidak boleh ada data kosong
            ' Cek kalau semua sudah diisi
            If MTxtNPWP.MaskCompleted = False Then
                MessageBox.Show("Please enter a valid 16-digit NPWP number.", "Validasi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                MTxtNPWP.Focus()
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub MTxtNPWP_MaskInputRejected(ByVal sender As Object, ByVal e As System.Windows.Forms.MaskInputRejectedEventArgs) Handles MTxtNPWP.MaskInputRejected
        Try
            MsgBox("Failed to modify data.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub MTxtNPWP_TypeValidationCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.TypeValidationEventArgs) Handles MTxtNPWP.TypeValidationCompleted
        Try
            MsgBox("Data has been successfully updated.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboTypeNPWP_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTypeNPWP.SelectedIndexChanged
        Try
            If Me.cboTypeNPWP.SelectedIndex = 0 Then
                Me.cboDocType.SelectedIndex = 0
                Me.cboDocType.Enabled = False
                Me.MTxtNPWP.Enabled = True
                Me.MTxtNPWP.Text = Me.NPWP
                Me.txtNIKPASS.Enabled = False
                Me.txtNIKPASS.Text = Me.NoNIKPASS
            Else
                Me.cboDocType.SelectedIndex = Me.Pilih_NIKPASS
                Me.cboDocType.Enabled = True
                Me.MTxtNPWP.Enabled = False
                Me.txtNIKPASS.Enabled = True
                Me.txtNIKPASS.Text = Me.NoNIKPASS
            End If
            Me.NPWPType = Me.cboTypeNPWP.SelectedIndex
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class