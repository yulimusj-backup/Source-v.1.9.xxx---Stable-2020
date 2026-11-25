Option Explicit On
Option Strict Off

Public Class frmLaporanVAT2025_EditSupplier
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public isEdit As Boolean = False
    Public PrimaryAcuan As Double = -1
    Public PrimaryTransaksi As Double = -1
    Public TypeTransaction As EnumTransactionType = EnumTransactionType.MasterSupplier
    Public TypeDocument As String = ""
    Public NPWP As String = ""
    Dim ObjSupp As New clsMasterSupplier
    Dim ObjKary As New ClsMasterKaryawan
    Dim DTA As New DseTrans_SettingNPWPTableAdapters.tblmaster_supplierTableAdapter
    Dim TBL As New DseTrans_SettingNPWP.tblmaster_supplierDataTable
    Dim DTAKaryawan As New DseTrans_SettingNPWPTableAdapters.tblmaster_karyawanTableAdapter
    Dim TBLKaryawan As New DseTrans_SettingNPWP.tblmaster_karyawanDataTable
    Dim DTAPurchase As New DseTrans_SettingNPWPTableAdapters.tbltrans_aopTableAdapter
    Dim TBLPurchase As New DseTrans_SettingNPWP.tbltrans_aopDataTable
    Dim DTAPurchaseReturn As New DseTrans_SettingNPWPTableAdapters.tbltrans_purchasereturnTableAdapter
    Dim TBLPurchaseReturn As New DseTrans_SettingNPWP.tbltrans_purchasereturnDataTable
    Dim NPWPType As Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP = Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP.Company

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

    Private Sub frmLaporanVAT2025_EditSupplier_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim NPWPSave As String
        Try
            If Me.isEdit = False Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar = True Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        NPWPSave = Me.MTxtNPWP.Text.Trim.Replace(" ", "")
                        'simpan ke master supplier
                        If Me.TypeTransaction = EnumTransactionType.MasterSupplier Then
                            DTA.Connection.ConnectionString = Core_Module.GetConstring
                            DTA.UpdateQuery_SupplierTaxData(Me.UserCode, Now, Now, Me.TxtAddressPajak1.Text, Me.TxtNPWP_Name.Text, _
                            Me.NPWPType, Me.TxtTKUID.Text, Me.TxtEmail.Text, NPWPSave, CDec(Me.PrimaryAcuan))
                        End If

                        'simpan ke master karyawan
                        If Me.TypeTransaction = EnumTransactionType.MasterEmployee Then
                            DTAKaryawan.Connection.ConnectionString = Core_Module.GetConstring
                            DTAKaryawan.UpdateQuery_TaxData(NPWPSave, Me.NPWPType, Me.UserCode, Now, Now, Me.TxtAddressPajak1.Text, _
                            Me.TxtEmail.Text, CDec(Me.PrimaryAcuan))
                        End If

                        'simpan ke transaksi AOP
                        If Me.TypeTransaction = EnumTransactionType.TransactionPurchase Then
                            DTAPurchase.Connection.ConnectionString = Core_Module.GetConstring
                            DTAPurchase.UpdateQuery_UpdateTaxData1(Me.UserCode, Now, NPWPSave, Me.TxtAddressPajak1.Text, CDec(Me.PrimaryTransaksi))
                        End If

                        'simpan ke transaksi Purchase return
                        If Me.TypeTransaction = EnumTransactionType.TransactionPurchaseReturn Then
                            DTAPurchaseReturn.Connection.ConnectionString = Core_Module.GetConstring
                            DTAPurchaseReturn.UpdateQuery_UpdateTaxData1(CDec(Me.PrimaryAcuan), NPWPSave, Me.TxtAddressPajak1.Text, Me.UserCode, Now, CDec(Me.PrimaryTransaksi))
                        End If

                        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
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

    Private Function ApakahDataBenar() As Boolean
        Try
            'Tidak boleh ada data kosong
            ' Cek kalau semua sudah diisi
            If MTxtNPWP.MaskCompleted = False Then
                MessageBox.Show("NPWP harus lengkap 16 digit!", "Validasi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                MTxtNPWP.Focus()
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Private Sub frmLaporanVAT2025_EditSupplier_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.txtPrimary.Text = Me.PrimaryAcuan
            Me.setData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub setData()
        Try
            If Me.TypeTransaction = EnumTransactionType.MasterEmployee Then
                Me.LoadDataKaryawan()
                Me.TxtAddressPajak1.Enabled = True
                Me.TxtTKUID.Enabled = False
                Me.TxtNPWP_Name.Enabled = False
            Else
                Me.LoadDataSupplier()
                Me.TxtAddressPajak1.Enabled = False
                Me.TxtTKUID.Enabled = True
                Me.TxtNPWP_Name.Enabled = True
            End If
            Me.MTxtNPWP.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function LoadDataSupplier() As Boolean
        Dim row As DseTrans_SettingNPWP.tblmaster_supplierRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, CDec(Me.PrimaryAcuan))
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.txtPrimary.Text = .Primary
                    Me.MTxtNPWP.Text = .NPWP_No.Replace(" ", "").Replace(".", "").Replace("-", "")
                    Me.TxtNPWP_Name.Text = .NPWP_Name
                    Me.TxtAddressPajak1.Text = .tax_address
                    Me.TxtTKUID.Text = .TKU_ID
                    Me.TxtEmail.Text = .Email
                    Me.NPWP = .NPWP_No.Replace(" ", "").Replace(".", "").Replace("-", "")
                    Me.cboTypeNPWP.SelectedIndex = .NPWP_Type
                End With
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function LoadDataKaryawan() As Boolean
        Dim row As DseTrans_SettingNPWP.tblmaster_karyawanRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, CDec(Me.PrimaryAcuan))
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.txtPrimary.Text = .Primary
                    Me.MTxtNPWP.Text = .npwp_no.Replace(" ", "").Replace(".", "").Replace("-", "")
                    Me.TxtNPWP_Name.Text = ""
                    Me.TxtAddressPajak1.Text = .npwp_address
                    Me.TxtTKUID.Text = ""
                    Me.TxtEmail.Text = .email
                    Me.NPWP = .npwp_no.Replace(" ", "").Replace(".", "").Replace("-", "")
                    Me.cboTypeNPWP.SelectedIndex = .NPWP_Type
                End With
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

#Region "Txt Focus"
    Private Sub MTxtNPWP_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MTxtNPWP.GotFocus
        Me.MTxtNPWP.BackColor = Me.BackColorAll
    End Sub

    Private Sub MTxtNPWP_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MTxtNPWP.LostFocus
        Me.MTxtNPWP.BackColor = Me.txtPrimary.BackColor
    End Sub

    Private Sub TxtAddressPajak1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddressPajak1.GotFocus
        Me.TxtAddressPajak1.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtAddressPajak1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddressPajak1.LostFocus
        Me.TxtAddressPajak1.BackColor = Me.txtPrimary.BackColor
    End Sub

    Private Sub TxtAddressPajak2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddressPajak2.GotFocus
        Me.TxtAddressPajak2.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtAddressPajak2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAddressPajak2.LostFocus
        Me.TxtAddressPajak2.BackColor = Me.txtPrimary.BackColor
    End Sub

    Private Sub TxtEmail_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmail.GotFocus
        Me.TxtEmail.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtEmail_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmail.LostFocus
        Me.TxtEmail.BackColor = Me.txtPrimary.BackColor
    End Sub

    Private Sub TxtNPWP_Name_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWP_Name.GotFocus
        Me.TxtNPWP_Name.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtNPWP_Name_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWP_Name.LostFocus
        Me.TxtNPWP_Name.BackColor = Me.txtPrimary.BackColor
    End Sub

    Private Sub TxtTaxCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxCode.GotFocus
        Me.TxtTaxCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtTaxCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxCode.LostFocus
        Me.TxtTaxCode.BackColor = Me.txtPrimary.BackColor
    End Sub

    Private Sub TxtTKUID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTKUID.GotFocus
        Me.TxtTKUID.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtTKUID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTKUID.LostFocus
        Me.TxtTKUID.BackColor = Me.txtPrimary.BackColor
    End Sub
#End Region

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
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
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    NPWPSave = Me.MTxtNPWP.Text.Trim.Replace(" ", "")

                    'simpan ke master supplier
                    If Me.TypeTransaction = EnumTransactionType.MasterSupplier Then
                        DTA.Connection.ConnectionString = Core_Module.GetConstring
                        DTA.UpdateQuery_SupplierTaxData(Me.UserCode, Now, Now, Me.TxtAddressPajak1.Text, Me.TxtNPWP_Name.Text, _
                        Me.NPWPType, Me.TxtTKUID.Text, Me.TxtEmail.Text, NPWPSave, CDec(Me.PrimaryAcuan))
                    End If

                    'simpan ke master karyawan
                    If Me.TypeTransaction = EnumTransactionType.MasterEmployee Then
                        DTAKaryawan.Connection.ConnectionString = Core_Module.GetConstring
                        DTAKaryawan.UpdateQuery_TaxData(NPWPSave, Me.NPWPType, Me.UserCode, Now, Now, Me.TxtAddressPajak1.Text, _
                        Me.TxtEmail.Text, CDec(Me.PrimaryAcuan))
                    End If

                    'simpan ke transaksi AOP
                    If Me.TypeTransaction = EnumTransactionType.TransactionPurchase Then
                        'simpan ke master supplier
                        DTA.Connection.ConnectionString = Core_Module.GetConstring
                        DTA.UpdateQuery_SupplierTaxData(Me.UserCode, Now, Now, Me.TxtAddressPajak1.Text, Me.TxtNPWP_Name.Text, _
                        Me.NPWPType, Me.TxtTKUID.Text, Me.TxtEmail.Text, NPWPSave, CDec(Me.PrimaryAcuan))

                        'DTAPurchase.Connection.ConnectionString = Core_Module.GetConstring
                        'DTAPurchase.UpdateQuery_UpdateTaxData1(Me.UserCode, Now, NPWPSave, Me.TxtAddressPajak1.Text, CDec(Me.PrimaryTransaksi))
                    End If

                    'simpan ke transaksi Purchase return
                    If Me.TypeTransaction = EnumTransactionType.TransactionPurchaseReturn Then
                        'simpan ke master supplier
                        DTA.Connection.ConnectionString = Core_Module.GetConstring
                        DTA.UpdateQuery_SupplierTaxData(Me.UserCode, Now, Now, Me.TxtAddressPajak1.Text, Me.TxtNPWP_Name.Text, _
                        Me.NPWPType, Me.TxtTKUID.Text, Me.TxtEmail.Text, NPWPSave, CDec(Me.PrimaryAcuan))

                        'DTAPurchaseReturn.Connection.ConnectionString = Core_Module.GetConstring
                        'DTAPurchaseReturn.UpdateQuery_UpdateTaxData1(CDec(Me.PrimaryAcuan), NPWPSave, Me.TxtAddressPajak1.Text, Me.UserCode, Now, CDec(Me.PrimaryTransaksi))
                    End If

                    Me.NPWP = Me.MTxtNPWP.Text
                    Me.isEdit = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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
End Class