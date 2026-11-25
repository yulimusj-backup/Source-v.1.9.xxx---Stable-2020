Option Explicit On
Option Strict Off

Imports System.Threading
Imports System.Globalization

Public Class FrmMasterKaryawan
    Public TBLKontrak As New DseMasterKaryawanTipe.tblmaster_karyawan_tipekontrakDataTable
    Dim RowKontrak As DseMasterKaryawanTipe.tblmaster_karyawan_tipekontrakRow
    Dim txtSelect As TextBox

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public Baru As Boolean
    Public ApakahPakaiCompetence As Boolean = True

    Dim ObjCekdata As New ClsCekData
    Dim ObjKaryawan As New ClsMasterKaryawan
    Dim ObjKaryawanType As New ClsMasterKaryawanTipe
    Dim ObjDept As New ClsMasterDepartement
    Dim ObjPost As New ClsMasterJobPosition
    Dim ObjKompetensi As New ClsMasterKompetensi
    Dim ObjItem As New ClsMasterKaryawan_Item
    Dim buffer As String = ""

    Dim mDgv As DataGridView
    Dim mEraseRowIndex As Integer
    Dim DTACountry As New DseMasterKaryawanTableAdapters.tblmaster_kode_negaraTableAdapter
    Dim TBLCountry As New DseMasterKaryawan.tblmaster_kode_negaraDataTable

    Public tblItemTipeKontrak As New DseMasterKaryawanTipe.tblmaster_karyawan_item_kontrakDataTable
    Public rowItemTipeKontrak As DseMasterKaryawanTipe.tblmaster_karyawan_item_kontrakRow
    Public tblItemAttch As New DseMasterKaryawanTipe.tblmaster_karyawan_item_kontrak_attachmentDataTable
    Public rowItemAttch As DseMasterKaryawanTipe.tblmaster_karyawan_item_kontrak_attachmentRow
    Dim cbo As DataGridViewComboBoxCell
    Dim ObjLicense As New ClsLicense

#Region "GotLostFocus+KeyPress"
    Private Sub TxtPersonal_Phone_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtPersonal_Phone.KeyPress
        'e.Handled = CumaAngka(Asc(e.KeyChar))
    End Sub

    Private Sub TxtEmployeeCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode.GotFocus
        Me.TxtEmployeeCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtEmployeeCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtEmployeeCode.KeyPress
        Me.BlokSimbol(Me.TxtEmployeeCode, e)
    End Sub

    Private Sub TxtEmployeeCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode.LostFocus
        Me.TxtEmployeeCode.BackColor = Me.TxtPrimaryEmployee.BackColor
    End Sub

    Private Sub TxtNIK_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNIK.GotFocus
        Me.TxtNIK.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtNIK_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtNIK.KeyPress
        Me.BlokSimbol(Me.TxtNIK, e)
    End Sub

    Private Sub TxtNIK_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNIK.LostFocus
        Me.TxtNIK.BackColor = Me.TxtPrimaryEmployee.BackColor
    End Sub

    Private Sub TxtDepartemen_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtDepartemen.KeyPress
        ''Me.BlokSimbol(Me.TxtDepartemen, e)
    End Sub

    Private Sub TxtDepartemen_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDepartemen.LostFocus
        Me.TxtDepartemen.BackColor = Me.TxtPrimaryEmployee.BackColor
    End Sub

    Private Sub TxtPositionCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPositionCode.GotFocus
        Me.TxtPositionCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtPositionCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtPositionCode.KeyPress
        ''Me.BlokSimbol(Me.TxtPositionCode, e)
    End Sub

    Private Sub TxtPositionCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPositionCode.LostFocus
        Me.TxtPositionCode.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtPersonal_BirthPlace_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPersonal_BirthPlace.GotFocus
        Me.TxtPersonal_BirthPlace.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtPersonal_BirthPlace_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtPersonal_BirthPlace.KeyPress
        Me.BlokSimbol(Me.TxtPersonal_BirthPlace, e)
    End Sub

    Private Sub TxtPersonal_BirthPlace_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPersonal_BirthPlace.LostFocus
        Me.TxtPersonal_BirthPlace.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtPersonal_Education_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPersonal_Education.GotFocus
        Me.TxtPersonal_Education.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtPersonal_Education_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtPersonal_Education.KeyPress
        Me.BlokSimbol(Me.TxtPersonal_Education, e)
    End Sub

    Private Sub TxtPersonal_Education_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPersonal_Education.LostFocus
        Me.TxtPersonal_Education.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtPersonal_Email_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPersonal_Email.GotFocus
        Me.TxtPersonal_Email.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtPersonal_Email_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtPersonal_Email.KeyPress
        Me.BlokSimbol(Me.TxtPersonal_Email, e)
    End Sub

    Private Sub TxtPersonal_Email_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPersonal_Email.Leave
        Try
            If Me.btnOK.Enabled = True Then
                Me.TxtPersonal_Email.Text = Me.TxtPersonal_Email.Text.Trim
                If Me.TxtPersonal_Email.Text <> "" Then
                    'pastikan formatnya adalah email
                    Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                    If Expression.IsMatch(Me.TxtPersonal_Email.Text) = False Then
                        MsgBox("The email address is not valid.", MsgBoxStyle.Information, "Information")
                        Me.TxtPersonal_Email.Focus()
                    End If
                    Expression = Nothing
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPersonal_Email_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPersonal_Email.LostFocus
        Me.TxtPersonal_Email.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtPersonal_Phone_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPersonal_Phone.GotFocus
        Me.TxtPersonal_Phone.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtPersonal_Phone_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPersonal_Phone.LostFocus
        Me.TxtPersonal_Phone.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtPersonal_Mother_Name_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPersonal_Mother_Name.GotFocus
        Me.TxtPersonal_Mother_Name.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtPersonal_Mother_Name_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtPersonal_Mother_Name.KeyPress
        Me.BlokSimbol(Me.TxtPersonal_Mother_Name, e)
    End Sub

    Private Sub TxtPersonal_Mother_Name_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPersonal_Mother_Name.LostFocus
        Me.TxtPersonal_Mother_Name.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtPersonal_Mother_Address_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPersonal_Mother_Address.GotFocus
        Me.TxtPersonal_Mother_Address.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtPersonal_Mother_Address_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtPersonal_Mother_Address.KeyPress
        Me.BlokSimbol(Me.TxtPersonal_Mother_Address, e)
    End Sub

    Private Sub TxtPersonal_Mother_Address_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPersonal_Mother_Address.LostFocus
        Me.TxtPersonal_Mother_Address.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtLegal_ID_Card_No_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLegal_ID_Card_No.GotFocus
        Me.TxtLegal_ID_Card_No.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtLegal_ID_Card_No_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtLegal_ID_Card_No.KeyPress
        Me.BlokSimbol(Me.TxtLegal_ID_Card_No, e)
    End Sub

    Private Sub TxtLegal_ID_Card_No_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLegal_ID_Card_No.LostFocus
        Me.TxtLegal_ID_Card_No.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtLegal_ID_Address_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLegal_ID_Address.GotFocus
        Me.TxtLegal_ID_Address.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtLegal_ID_Address_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtLegal_ID_Address.KeyPress
        Me.BlokSimbol(Me.TxtLegal_ID_Card_No, e)
    End Sub

    Private Sub TxtLegal_ID_Address_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLegal_ID_Address.LostFocus
        Me.TxtLegal_ID_Address.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtLegal_Current_Address_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLegal_Current_Address.GotFocus
        Me.TxtLegal_Current_Address.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtLegal_Current_Address_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtLegal_Current_Address.KeyPress
        Me.BlokSimbol(Me.TxtLegal_Current_Address, e)
    End Sub

    Private Sub TxtLegal_Current_Address_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLegal_Current_Address.LostFocus
        Me.TxtLegal_Current_Address.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtPersonal_NPWP_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPersonal_NPWP.GotFocus
        Me.TxtPersonal_NPWP.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtPersonal_NPWP_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtPersonal_NPWP.KeyPress
        Me.BlokSimbol(Me.TxtPersonal_NPWP, e)
    End Sub

    Private Sub TxtPersonal_NPWP_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPersonal_NPWP.LostFocus
        Me.TxtPersonal_NPWP.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtLegal_NPWP_Address_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLegal_NPWP_Address.GotFocus
        Me.TxtLegal_NPWP_Address.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtLegal_NPWP_Address_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtLegal_NPWP_Address.KeyPress
        Me.BlokSimbol(Me.TxtLegal_NPWP_Address, e)
    End Sub

    Private Sub TxtLegal_NPWP_Address_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLegal_NPWP_Address.LostFocus
        Me.TxtLegal_NPWP_Address.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtLegal_Astek_No_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLegal_Astek_No.GotFocus
        Me.TxtLegal_Astek_No.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtLegal_Astek_No_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtLegal_Astek_No.KeyPress
        Me.BlokSimbol(Me.TxtLegal_Astek_No, e)
    End Sub

    Private Sub TxtLegal_Astek_No_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLegal_Astek_No.LostFocus
        Me.TxtLegal_Astek_No.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtLegal_KPJ_No_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLegal_KPJ_No.GotFocus
        Me.TxtLegal_KPJ_No.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtLegal_KPJ_No_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtLegal_KPJ_No.KeyPress
        Me.BlokSimbol(Me.TxtLegal_KPJ_No, e)
    End Sub

    Private Sub TxtLegal_KPJ_No_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLegal_KPJ_No.LostFocus
        Me.TxtLegal_KPJ_No.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtBank_Acc_Holder_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBank_Acc_Holder.GotFocus
        Me.TxtBank_Acc_Holder.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtBank_Acc_Holder_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBank_Acc_Holder.KeyPress
        Me.BlokSimbol(Me.TxtBank_Acc_Holder, e)
    End Sub

    Private Sub TxtBank_Acc_Holder_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBank_Acc_Holder.LostFocus
        Me.TxtBank_Acc_Holder.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtBank_Acc_No_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBank_Acc_No.GotFocus
        Me.TxtBank_Acc_No.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtBank_Acc_No_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBank_Acc_No.KeyPress
        Me.BlokSimbol(Me.TxtBank_Acc_No, e)
    End Sub

    Private Sub TxtBank_Acc_No_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBank_Acc_No.LostFocus
        Me.TxtBank_Acc_No.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtBank_Name_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBank_Name.GotFocus
        Me.TxtBank_Name.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtBank_Name_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBank_Name.KeyPress
        Me.BlokSimbol(Me.TxtBank_Name, e)
    End Sub

    Private Sub TxtBank_Name_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBank_Name.LostFocus
        Me.TxtBank_Name.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtBank_Branch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBank_Branch.GotFocus
        Me.TxtBank_Branch.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtBank_Branch_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBank_Branch.KeyPress
        Me.BlokSimbol(Me.TxtBank_Branch, e)
    End Sub

    Private Sub TxtBank_Branch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBank_Branch.LostFocus
        Me.TxtBank_Branch.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtNotes_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNotes.GotFocus
        Me.TxtNotes.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtNotes_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtNotes.KeyPress
        Me.BlokSimbol(Me.TxtNotes, e)
    End Sub

    Private Sub TxtNotes_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNotes.LostFocus
        Me.TxtNotes.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtDepartemen_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDepartemen.GotFocus
        Me.TxtDepartemen.BackColor = Me.BackColorAll
    End Sub

    'aul tambahan got focus/lost focus txtMedicalCentre dan txtBank_Code
    Private Sub TxtMedicalCenter_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMedicalCenter.GotFocus
        Me.TxtMedicalCenter.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtMedicalCenter_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtMedicalCenter.KeyPress
        Me.BlokSimbol(Me.TxtMedicalCenter, e)
    End Sub

    Private Sub TxtMedicalCenter_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMedicalCenter.LostFocus
        Me.TxtMedicalCenter.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

    Private Sub TxtBank_Code_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBank_Code.GotFocus
        Me.TxtBank_Code.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtBank_Code_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBank_Code.KeyPress
        Me.BlokSimbol(Me.TxtBank_Code, e)
    End Sub

    Private Sub TxtBank_Code_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBank_Code.LostFocus
        Me.TxtBank_Code.BackColor = Me.TxtPrimaryDepartemen.BackColor
    End Sub

#End Region

    Private Sub CboCountryDesc_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboCountryDesc.SelectedIndexChanged
        Dim row As DseMasterKaryawan.tblmaster_kode_negaraRow
        Try
            If Me.CboForeignEmployee.Enabled = True Then
                If Me.CboCountryDesc.SelectedIndex = -1 Then
                    Me.TxtPrimaryCountryCode.Text = "-1"
                Else
                    row = Me.TBLCountry.DefaultView(Me.CboCountryDesc.SelectedIndex).Row
                    Me.TxtPrimaryCountryCode.Text = row.Primary
                    Me.TxtPersonal_Nationality.Text = row.Nama_Negara
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboForeignEmployee_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboForeignEmployee.SelectedIndexChanged
        Try
            If Me.CboForeignEmployee.Enabled = True Then
                If Me.CboForeignEmployee.SelectedIndex = 0 Then
                    Me.CboCountryDesc.Enabled = False
                    Me.TxtPersonal_Nationality.Text = "Indonesia"
                Else
                    Me.CboCountryDesc.Enabled = True
                    Me.TxtPersonal_Nationality.Text = "Afganistan"
                End If
                Me.CboCountryDesc.SelectedIndex = 0
                Me.TxtPrimaryCountryCode.Text = "1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub UpdateStandarValue()
        Dim rowDgv As DataGridViewRow
        Dim i As Integer = 1
        Try
            If Me.DGVKompetensi.Rows.Count > 1 Then
                For Each rowDgv In Me.DGVKompetensi.Rows
                    If i <> Me.DGVKompetensi.Rows.Count Then
                        If Me.ObjPost.GetJmlItem(Me.TxtPrimaryPosition.Text, rowDgv.Cells(Me.ColCompetence_Primary.Index).Value) > 0 Then
                            rowDgv.Cells(Me.ColCompetence_StandarValue.Index).Value = _
                            Me.ObjPost.GetValue_CustomFormatted(Me.TxtPrimaryPosition.Text, rowDgv.Cells(Me.ColCompetence_Primary.Index).Value _
                            , Me.ObjCekdata, rowDgv.Cells(Me.ColCempetence_DataType.Index).Value)
                        Else
                            rowDgv.Cells(Me.ColCompetence_StandarValue.Index).Value = _
                            Me.ObjKompetensi.GetValue_CustomFormatted(rowDgv.Cells(Me.ColCompetence_Primary.Index).Value _
                            , Me.ObjCekdata, rowDgv.Cells(Me.ColCempetence_DataType.Index).Value)
                        End If
                    End If
                    i = i + 1
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadKompetensi(ByVal vPrimaryJob As Double, ByVal vApakahPakaiCompetence As Boolean)
        Dim hasil As MsgBoxResult
        Dim Dta As New DseMasterJobPositionTableAdapters.tblmaster_jobposition_item_kompetensi1TableAdapter
        Dim Tbl As New DseMasterJobPosition.tblmaster_jobposition_item_kompetensi1DataTable
        Dim row As DseMasterJobPosition.tblmaster_jobposition_item_kompetensi1Row
        Dim rowDgv As DataGridViewRow
        Dim DefPrec As Integer = 0

        Try
            If Me.DGVKompetensi.Rows.Count <= 1 Then
                If vApakahPakaiCompetence = True Then
                    If Me.ObjPost.ApakahMemilikiKompetensi(Me.TxtPrimaryPosition.Text) = True Then
                        hasil = MsgBox("Do you want to load competence from this job position ?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Dta.Fill(Tbl, vPrimaryJob)
                            For Each row In Tbl.Rows
                                Me.DGVKompetensi.Rows.Add()
                                If Me.DGVKompetensi.Rows.Count > 1 Then
                                    rowDgv = Me.DGVKompetensi.Rows(Me.DGVKompetensi.RowCount - 2)
                                Else
                                    rowDgv = Me.DGVKompetensi.Rows(Me.DGVKompetensi.RowCount)
                                End If
                                rowDgv.Cells(Me.ColCompetence_Code.Index).Value = row.Competence_Code
                                rowDgv.Cells(Me.ColCempetence_Description.Index).Value = row.Description
                                rowDgv.Cells(Me.ColCompetence_StandarValue.Index).Value = Me.ObjPost.GetValue_CustomFormatted(row.primary_job_position, row.primary_kompetensi, _
                                Me.ObjCekdata, row.Data_Type)
                                rowDgv.Cells(Me.ColCompetence_Value.Index).Value = ""
                                If row.periodic_update = 1 Then
                                    rowDgv.Cells(Me.ColCompetence_NextCertificationDate.Index).Value = Format(Now, "dd MMM yyyy")
                                Else
                                    rowDgv.Cells(Me.ColCompetence_NextCertificationDate.Index).Value = "-"
                                End If
                                rowDgv.Cells(Me.ColCompetence_Remarks.Index).Value = ""
                                rowDgv.Cells(Me.ColCompetence_Primary.Index).Value = row.primary_kompetensi
                                rowDgv.Cells(Me.ColCempetence_DataType.Index).Value = row.Data_Type
                            Next
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemInput(ByVal vPrimary As Double) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.DGVKompetensi.Rows
                If rowDGV.Cells(Me.ColCompetence_Primary.Index).Value = vPrimary Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Function BlokSimbol(ByVal obj As TextBox, ByVal e As Windows.Forms.KeyPressEventArgs) As Boolean
        Try
            If e.KeyChar = "'" Or _
            e.KeyChar = "%" Or _
            e.KeyChar = "*" Or _
            e.KeyChar = "#" Then
                obj.ReadOnly = True
                obj.BackColor = Me.BackColorAll
            Else
                obj.ReadOnly = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function CumaAngka(ByVal kode As String, Optional ByVal spasi As Boolean = False) As Boolean
        Select Case kode
            Case 48 To 57
                Return False
            Case 8
                Return False
            Case 45
                Return True
            Case 32
                If spasi Then
                    Return False
                Else
                    Return True
                End If
            Case Else
                Return True
        End Select
    End Function

    Private Sub Aktifkandetail()
        Try
            Me.ObjCekdata.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCekdata.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCekdata.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCekdata.KendalikanControl(Me.TabPage1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCekdata.KendalikanControl(Me.TabPage2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCekdata.KendalikanControl(Me.TabPage3.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCekdata.KendalikanControl(Me.TabPage4.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCekdata.KendalikanControl(Me.TabPage7.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ChkOverrideClockIn.Enabled = True
            Me.ChkOverrideClockOut.Enabled = True
            If Me.ChkOverrideClockIn.Checked = True Then
                Me.ChkDoesnNotClockIn.Enabled = True
            Else
                Me.ChkDoesnNotClockIn.Enabled = False
            End If
            If Me.ChkOverrideClockOut.Checked = True Then
                Me.ChkDoesnNotClockOut.Enabled = True
            Else
                Me.ChkDoesnNotClockOut.Enabled = False
            End If

            Me.ChkOverrideOTOrder.Enabled = True
            Me.ChkOverrideEarlyOT.Enabled = True
            Me.ChkOverrideLate.Enabled = True
            Me.ChkOverrideEarlyOut.Enabled = True
            Me.ChkOverrideOT.Enabled = True
            If Me.ChkOverrideEarlyOT.Checked = True Then
                Me.CboIsCalculateEarlyOT.Enabled = True
                If Me.CboIsCalculateEarlyOT.SelectedIndex = 0 Then
                    Me.txtCalculateAfterEarlyOT.Enabled = True
                    Me.cboRoundingTypeEarlyOT.Enabled = True
                    If Me.cboRoundingTypeEarlyOT.SelectedIndex <> 0 Then
                        Me.txtRoundingEarlyOT.Enabled = True
                    End If
                    Me.txtMaxValueEarlyOT.Enabled = True
                    Me.chkOverrideCalculateEarlyOTTolerance.Enabled = True
                Else
                    Me.chkOverrideCalculateEarlyOTTolerance.Enabled = False
                End If
            ElseIf Me.ChkOverrideEarlyOT.Checked = False Then
                Me.CboIsCalculateEarlyOT.Enabled = False
                Me.txtCalculateAfterEarlyOT.Enabled = False
                Me.cboRoundingTypeEarlyOT.Enabled = False
                Me.txtRoundingEarlyOT.Enabled = False
                Me.txtMaxValueEarlyOT.Enabled = False
                Me.chkOverrideCalculateEarlyOTTolerance.Enabled = False
            End If
            If Me.ChkOverrideLate.Checked = True Then
                Me.CboIsCalculateLate.Enabled = True
                If Me.CboIsCalculateLate.SelectedIndex = 0 Then
                    Me.txtCalculateAfterLate.Enabled = True
                    Me.cboRoundingTypeLate.Enabled = True
                    If Me.cboRoundingTypeLate.SelectedIndex <> 0 Then
                        Me.txtRoundingLate.Enabled = True
                    End If
                    Me.txtMaxValueLate.Enabled = True
                    Me.chkOverrideCalculateLateTolerance.Enabled = True
                Else
                    Me.chkOverrideCalculateLateTolerance.Enabled = False
                End If
            ElseIf Me.ChkOverrideLate.Checked = False Then
                Me.CboIsCalculateLate.Enabled = False
                Me.txtCalculateAfterLate.Enabled = False
                Me.cboRoundingTypeLate.Enabled = False
                Me.txtRoundingLate.Enabled = False
                Me.txtMaxValueLate.Enabled = False
                Me.chkOverrideCalculateLateTolerance.Enabled = False
            End If
            If Me.ChkOverrideEarlyOut.Checked = True Then
                Me.CboIsCalculateEarlyOut.Enabled = True
                If Me.CboIsCalculateEarlyOut.SelectedIndex = 0 Then
                    Me.txtCalculateAfterEarlyOut.Enabled = True
                    Me.cboRoundingTypeEarlyOut.Enabled = True
                    If Me.cboRoundingTypeEarlyOut.SelectedIndex <> 0 Then
                        Me.txtRoundingEarlyOut.Enabled = True
                    End If
                    Me.txtMaxValueEarlyOut.Enabled = True
                    Me.chkOverrideCalculateEarlyOutTolerance.Enabled = True
                Else
                    Me.chkOverrideCalculateEarlyOutTolerance.Enabled = False
                End If
            ElseIf Me.ChkOverrideEarlyOut.Checked = False Then
                Me.CboIsCalculateEarlyOut.Enabled = False
                Me.txtCalculateAfterEarlyOut.Enabled = False
                Me.cboRoundingTypeEarlyOut.Enabled = False
                Me.txtRoundingEarlyOut.Enabled = False
                Me.txtMaxValueEarlyOut.Enabled = False
                Me.chkOverrideCalculateEarlyOutTolerance.Enabled = False
            End If
            If Me.ChkOverrideOT.Checked = True Then
                Me.CboIsCalculateOT.Enabled = True
                If Me.CboIsCalculateOT.SelectedIndex = 0 Then
                    Me.txtCalculateAfterOT.Enabled = True
                    Me.cboRoundingTypeOT.Enabled = True
                    If Me.cboRoundingTypeOT.SelectedIndex <> 0 Then
                        Me.txtRoundingOT.Enabled = True
                    End If
                    Me.txtMaxValueOT.Enabled = True
                    Me.chkOverrideCalculateOTTolerance.Enabled = True
                Else
                    Me.chkOverrideCalculateOTTolerance.Enabled = False
                End If
            ElseIf Me.ChkOverrideOT.Checked = False Then
                Me.CboIsCalculateOT.Enabled = False
                Me.txtCalculateAfterOT.Enabled = False
                Me.cboRoundingTypeOT.Enabled = False
                Me.txtRoundingOT.Enabled = False
                Me.txtMaxValueOT.Enabled = False
                Me.chkOverrideCalculateOTTolerance.Enabled = False
            End If
            Me.dgvFamily.AllowUserToAddRows = True
            Me.dgvFamily.AllowUserToDeleteRows = True
            Me.dgvFamily.AllowUserToResizeColumns = True
            Me.dgvFamily.AllowUserToResizeRows = True
            Me.dgvFamily.ReadOnly = False
            Me.ChkNoNPWP.Enabled = True
            If Me.ChkNoNPWP.Checked = False Then
                Me.TxtLegal_NPWP_Address.Enabled = False
                Me.TxtPersonal_NPWP.Enabled = False
            End If
            Me.DGVKompetensi.AllowUserToAddRows = True
            Me.DGVKompetensi.ReadOnly = False
            Me.DGVKompetensi.AllowUserToDeleteRows = True
            Me.CancelButton = Me.btnCancel
            Me.CboForeignEmployee.Enabled = True
            If Me.CboForeignEmployee.SelectedIndex = 0 Then
                Me.CboCountryDesc.Enabled = False
            Else
                Me.CboCountryDesc.Enabled = True
            End If
            Me.TxtPersonal_Nationality.Enabled = False
            If Me.chkUseOvertimeOrder.Checked = False Then
                Me.chkSendNotifEmailOvertimeOrder.Checked = False
                Me.chkSendNotifEmailOvertimeOrder.Enabled = False
            Else
                Me.chkSendNotifEmailOvertimeOrder.Enabled = True
            End If
            'Me.chkUseOvertimeOrder.Enabled = True
            'Me.ObjCekdata.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False, True, False)
            'Me.ObjCekdata.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            'Me.ObjCekdata.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            'Me.ObjCekdata.KendalikanControl(Me.TabPage1.Controls, True, True, True, True, True, True, True, True, True, True)
            'Me.ObjCekdata.KendalikanControl(Me.TabPage2.Controls, True, True, True, True, True, True, True, True, True, True)
            'Me.ObjCekdata.KendalikanControl(Me.TabPage3.Controls, True, True, True, True, True, True, True, True, True, True)
            'Me.ObjCekdata.KendalikanControl(Me.TabPage4.Controls, True, True, True, True, True, True, True, True, True, True)
            'Me.ObjCekdata.KendalikanControl(Me.TabPage7.Controls, True, True, True, True, True, True, True, True, True, True)
            'Me.ChkOverrideClockIn.Enabled = True
            'Me.ChkOverrideClockOut.Enabled = True
            'If Me.ChkOverrideClockIn.Checked = True Then
            '    Me.ChkDoesnNotClockIn.Enabled = True
            'Else
            '    Me.ChkDoesnNotClockIn.Enabled = False
            'End If
            'If Me.ChkOverrideClockOut.Checked = True Then
            '    Me.ChkDoesnNotClockOut.Enabled = True
            'Else
            '    Me.ChkDoesnNotClockOut.Enabled = False
            'End If

            'Me.ChkOverrideEarlyOT.Enabled = True
            'Me.ChkOverrideLate.Enabled = True
            'Me.ChkOverrideEarlyOut.Enabled = True
            'Me.ChkOverrideOT.Enabled = True
            'If Me.ChkOverrideEarlyOT.Checked = True Then
            '    Me.CboIsCalculateEarlyOT.Enabled = True
            '    If Me.CboIsCalculateEarlyOT.SelectedIndex = 0 Then
            '        Me.txtCalculateAfterEarlyOT.Enabled = True
            '        Me.cboRoundingTypeEarlyOT.Enabled = True
            '        If Me.cboRoundingTypeEarlyOT.SelectedIndex <> 0 Then
            '            Me.txtRoundingEarlyOT.Enabled = True
            '        End If
            '        Me.txtMaxValueEarlyOT.Enabled = True
            '    End If
            '    Me.chkOverrideCalculateEarlyOTTolerance.Enabled = True
            'ElseIf Me.ChkOverrideEarlyOT.Checked = False Then
            '    Me.CboIsCalculateEarlyOT.Enabled = False
            '    Me.txtCalculateAfterEarlyOT.Enabled = False
            '    Me.cboRoundingTypeEarlyOT.Enabled = False
            '    Me.txtRoundingEarlyOT.Enabled = False
            '    Me.txtMaxValueEarlyOT.Enabled = False
            '    Me.chkOverrideCalculateEarlyOTTolerance.Enabled = False
            'End If
            'If Me.ChkOverrideLate.Checked = True Then
            '    Me.CboIsCalculateLate.Enabled = True
            '    If Me.CboIsCalculateLate.SelectedIndex = 0 Then
            '        Me.txtCalculateAfterLate.Enabled = True
            '        Me.cboRoundingTypeLate.Enabled = True
            '        If Me.cboRoundingTypeLate.SelectedIndex <> 0 Then
            '            Me.txtRoundingLate.Enabled = True
            '        End If
            '        Me.txtMaxValueLate.Enabled = True
            '    End If
            '    Me.chkOverrideCalculateLateTolerance.Enabled = True
            'ElseIf Me.ChkOverrideLate.Checked = False Then
            '    Me.CboIsCalculateLate.Enabled = False
            '    Me.txtCalculateAfterLate.Enabled = False
            '    Me.cboRoundingTypeLate.Enabled = False
            '    Me.txtRoundingLate.Enabled = False
            '    Me.txtMaxValueLate.Enabled = False
            '    Me.chkOverrideCalculateLateTolerance.Enabled = False
            'End If
            'If Me.ChkOverrideEarlyOut.Checked = True Then
            '    Me.CboIsCalculateEarlyOut.Enabled = True
            '    If Me.CboIsCalculateEarlyOut.SelectedIndex = 0 Then
            '        Me.txtCalculateAfterEarlyOut.Enabled = True
            '        Me.cboRoundingTypeEarlyOut.Enabled = True
            '        If Me.cboRoundingTypeEarlyOut.SelectedIndex <> 0 Then
            '            Me.txtRoundingEarlyOut.Enabled = True
            '        End If
            '        Me.txtMaxValueEarlyOut.Enabled = True
            '    End If
            '    Me.chkOverrideCalculateEarlyOutTolerance.Enabled = True
            'ElseIf Me.ChkOverrideEarlyOut.Checked = False Then
            '    Me.CboIsCalculateEarlyOut.Enabled = False
            '    Me.txtCalculateAfterEarlyOut.Enabled = False
            '    Me.cboRoundingTypeEarlyOut.Enabled = False
            '    Me.txtRoundingEarlyOut.Enabled = False
            '    Me.txtMaxValueEarlyOut.Enabled = False
            '    Me.chkOverrideCalculateEarlyOutTolerance.Enabled = False
            'End If
            'If Me.ChkOverrideOT.Checked = True Then
            '    Me.CboIsCalculateOT.Enabled = True
            '    If Me.CboIsCalculateOT.SelectedIndex = 0 Then
            '        Me.txtCalculateAfterOT.Enabled = True
            '        Me.cboRoundingTypeOT.Enabled = True
            '        If Me.cboRoundingTypeOT.SelectedIndex <> 0 Then
            '            Me.txtRoundingOT.Enabled = True
            '        End If
            '        Me.txtMaxValueOT.Enabled = True
            '    End If
            '    Me.chkOverrideCalculateOTTolerance.Enabled = True
            'ElseIf Me.ChkOverrideOT.Checked = False Then
            '    Me.CboIsCalculateOT.Enabled = False
            '    Me.txtCalculateAfterOT.Enabled = False
            '    Me.cboRoundingTypeOT.Enabled = False
            '    Me.txtRoundingOT.Enabled = False
            '    Me.txtMaxValueOT.Enabled = False
            '    Me.chkOverrideCalculateOTTolerance.Enabled = False
            'End If
            'Me.dgvFamily.AllowUserToAddRows = True
            'Me.dgvFamily.AllowUserToDeleteRows = True
            'Me.dgvFamily.AllowUserToResizeColumns = True
            'Me.dgvFamily.AllowUserToResizeRows = True
            'Me.dgvFamily.ReadOnly = False
            'Me.ChkNoNPWP.Enabled = True
            'If Me.ChkNoNPWP.Checked = False Then
            '    Me.TxtLegal_NPWP_Address.Enabled = False
            '    Me.TxtPersonal_NPWP.Enabled = False
            'End If
            'Me.DGVKompetensi.AllowUserToAddRows = True
            'Me.DGVKompetensi.ReadOnly = False
            'Me.DGVKompetensi.AllowUserToDeleteRows = True
            'Me.CancelButton = Me.btnCancel
            'Me.CboForeignEmployee.Enabled = True
            'If Me.CboForeignEmployee.SelectedIndex = 0 Then
            '    Me.CboCountryDesc.Enabled = False
            'Else
            '    Me.CboCountryDesc.Enabled = True
            'End If
            'Me.TxtPersonal_Nationality.Enabled = False

            Me.DGVKontrak.AllowUserToAddRows = True
            Me.DGVKontrak.AllowUserToDeleteRows = True
            Me.DGVKontrak.AllowUserToResizeColumns = True
            Me.DGVKontrak.AllowUserToResizeRows = True
            Me.DGVKontrak.ReadOnly = False

            Me.DGVAttanchmentKontrak.AllowUserToAddRows = True
            Me.DGVAttanchmentKontrak.AllowUserToDeleteRows = True
            Me.DGVAttanchmentKontrak.AllowUserToResizeColumns = True
            Me.DGVAttanchmentKontrak.AllowUserToResizeRows = True
            Me.DGVAttanchmentKontrak.ReadOnly = False

            Me.DGVAttanchment.AllowUserToAddRows = True
            Me.DGVAttanchment.AllowUserToDeleteRows = True
            Me.DGVAttanchment.AllowUserToResizeColumns = True
            Me.DGVAttanchment.AllowUserToResizeRows = True
            Me.DGVAttanchment.ReadOnly = False

            Me.SetUseOvertimeOrder()
            Me.llbCopy.Enabled = False

            Me.chkloses_tax_obligations.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCekdata.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCekdata.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCekdata.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCekdata.KendalikanControl(Me.TabPage1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCekdata.KendalikanControl(Me.TabPage2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCekdata.KendalikanControl(Me.TabPage3.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCekdata.KendalikanControl(Me.TabPage4.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCekdata.KendalikanControl(Me.TabPage7.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ChkOverrideClockIn.Enabled = False
            Me.ChkOverrideClockOut.Enabled = False
            Me.ChkDoesnNotClockIn.Enabled = False
            Me.ChkDoesnNotClockOut.Enabled = False

            Me.ChkOverrideOTOrder.Enabled = False
            Me.chkUseOvertimeOrder.Enabled = False
            Me.chkSendNotifEmailOvertimeOrder.Enabled = False
            Me.ChkOverrideEarlyOT.Enabled = False
            Me.ChkOverrideLate.Enabled = False
            Me.ChkOverrideEarlyOut.Enabled = False
            Me.ChkOverrideOT.Enabled = False
            Me.CboIsCalculateEarlyOT.Enabled = False
            Me.CboIsCalculateLate.Enabled = False
            Me.CboIsCalculateEarlyOut.Enabled = False
            Me.CboIsCalculateOT.Enabled = False
            Me.txtCalculateAfterEarlyOT.Enabled = False
            Me.txtCalculateAfterLate.Enabled = False
            Me.txtCalculateAfterEarlyOut.Enabled = False
            Me.txtCalculateAfterOT.Enabled = False
            Me.cboRoundingTypeEarlyOT.Enabled = False
            Me.cboRoundingTypeLate.Enabled = False
            Me.cboRoundingTypeEarlyOut.Enabled = False
            Me.cboRoundingTypeOT.Enabled = False
            Me.txtRoundingEarlyOT.Enabled = False
            Me.txtRoundingLate.Enabled = False
            Me.txtRoundingEarlyOut.Enabled = False
            Me.txtRoundingOT.Enabled = False
            Me.txtMaxValueEarlyOT.Enabled = False
            Me.txtMaxValueLate.Enabled = False
            Me.txtMaxValueEarlyOut.Enabled = False
            Me.txtMaxValueOT.Enabled = False

            Me.dgvFamily.AllowUserToAddRows = False
            Me.dgvFamily.AllowUserToDeleteRows = False
            Me.dgvFamily.AllowUserToResizeColumns = False
            Me.dgvFamily.AllowUserToResizeRows = False
            Me.DGVKompetensi.AllowUserToAddRows = False
            Me.DGVKompetensi.AllowUserToDeleteRows = False
            Me.DGVKompetensi.AllowUserToResizeColumns = False
            Me.DGVKompetensi.AllowUserToResizeRows = False
            Me.DGVKontrak.AllowUserToAddRows = False
            Me.DGVKontrak.AllowUserToDeleteRows = False
            Me.DGVKontrak.AllowUserToResizeColumns = False
            Me.DGVKontrak.AllowUserToResizeRows = False
            Me.DGVAttanchmentKontrak.AllowUserToAddRows = False
            Me.DGVAttanchmentKontrak.AllowUserToDeleteRows = False
            Me.DGVAttanchmentKontrak.AllowUserToResizeColumns = False
            Me.DGVAttanchmentKontrak.AllowUserToResizeRows = False
            Me.DGVAttanchment.AllowUserToAddRows = False
            Me.DGVAttanchment.AllowUserToDeleteRows = False
            Me.DGVAttanchment.AllowUserToResizeColumns = False
            Me.DGVAttanchment.AllowUserToResizeRows = False

            Me.dgvFamily.ReadOnly = True
            Me.DGVKompetensi.ReadOnly = True
            Me.DGVKontrak.ReadOnly = True
            Me.DGVAttanchmentKontrak.ReadOnly = True
            Me.DGVAttanchment.ReadOnly = True
            Me.ChkNoNPWP.Enabled = False
            Me.TxtPersonal_NPWP.Enabled = False
            Me.TxtLegal_NPWP_Address.Enabled = False
            Me.CancelButton = Me.btnExit
            Me.CboForeignEmployee.Enabled = False
            Me.CboCountryDesc.Enabled = False

            Me.chkOverrideCalculateEarlyOTTolerance.Enabled = False
            Me.chkOverrideCalculateLateTolerance.Enabled = False
            Me.chkOverrideCalculateEarlyOutTolerance.Enabled = False
            Me.chkOverrideCalculateOTTolerance.Enabled = False
            Me.chkUseOvertimeOrder.Enabled = False
            Me.chkSendNotifEmailOvertimeOrder.Enabled = False
            Me.llbCopy.Enabled = True
            Me.chkloses_tax_obligations.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCekdata.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            Me.ObjCekdata.KosongkanControl(Me.TabPage1.Controls, True, False, False)
            Me.ObjCekdata.KosongkanControl(Me.TabPage2.Controls, True, False, False)
            Me.ObjCekdata.KosongkanControl(Me.TabPage3.Controls, True, False, False)
            Me.ObjCekdata.KosongkanControl(Me.TabPage4.Controls, True, False, False)
            Me.ObjCekdata.KosongkanControl(Me.GroupBox2.Controls, True, True, True)
            Me.ObjCekdata.KosongkanControl(Me.TabPage7.Controls, True, False, False)
            Me.ChkOverrideClockIn.Checked = False
            Me.ChkOverrideClockOut.Checked = False
            Me.ChkDoesnNotClockIn.Checked = False
            Me.ChkDoesnNotClockOut.Checked = False

            Me.PictureBox1.Image = Nothing
            Me.AturPosisiKomponenGambar()
            Me.CboEmployeeStatus.SelectedIndex = ClsMasterKaryawan.EnumEmployStatus.Probation
            Me.cboStatus.SelectedIndex = ClsMasterKaryawan.EnumAktif.Active
            Me.CboPersonal_Sex.SelectedIndex = ClsMasterKaryawan.EnumSex.Male
            Me.CboPersonal_Religion.SelectedIndex = ClsMasterKaryawan.EnumReligion.Islam
            Me.CboPersonal_BloodType.SelectedIndex = ClsMasterKaryawan.EnumBloodType.A
            Me.CboPersonal_MartialStatus.SelectedIndex = ClsMasterKaryawan.EnumMartialStatus._Single
            Me.CboLegal_Category.SelectedIndex = ClsMasterKaryawan.GolonganPajak.TK0
            Me.TxtPrimaryDepartemen.Text = -1
            Me.TxtPrimaryEmployee.Text = -1
            Me.TxtPrimaryPosition.Text = -1
            Me.TextBox11.Text = -1
            Me.ToolStripStatusLabel1.Text = "Created By: "
            Me.ToolStripStatusLabel2.Text = "Created Date: "
            Me.ToolStripStatusLabel3.Text = "Modified By: "
            Me.ToolStripStatusLabel4.Text = "Modified Date: "
            Me.TxtPersonal_NPWP.Text = "00.000.000.0-000.000"
            Me.ChkNoNPWP.Checked = False
            Me.dgvFamily.Rows.Clear()
            Me.DGVKompetensi.Rows.Clear()
            Me.DGVKontrak.Rows.Clear()
            Me.DGVAttanchmentKontrak.Rows.Clear()
            Me.DGVAttanchment.Rows.Clear()
            '--------------------------------------
            Me.CboForeignEmployee.SelectedIndex = 0
            Me.TxtPrimaryCountryCode.Text = "-1"
            Me.ObjKaryawan.IsiTBLKodeNegara(Me.CboCountryDesc, Me.TBLCountry, Me.DTACountry, Me)
            If Me.CboForeignEmployee.SelectedIndex = 0 Then
                Me.TxtPersonal_Nationality.Text = "Indonesia"
            End If

            Me.DtpResignDate.Value = Now

            Me.ChkOverrideOTOrder.Checked = False
            Me.ChkOverrideEarlyOT.Checked = False
            Me.ChkOverrideLate.Checked = False
            Me.ChkOverrideEarlyOut.Checked = False
            Me.ChkOverrideOT.Checked = False
            Me.CboIsCalculateEarlyOT.SelectedIndex = 0
            Me.CboIsCalculateLate.SelectedIndex = 0
            Me.CboIsCalculateEarlyOut.SelectedIndex = 0
            Me.CboIsCalculateOT.SelectedIndex = 0
            Me.txtCalculateAfterEarlyOT.Text = "00.00"
            Me.txtCalculateAfterLate.Text = "00.00"
            Me.txtCalculateAfterEarlyOut.Text = "00.00"
            Me.txtCalculateAfterOT.Text = "00.00"
            Me.cboRoundingTypeEarlyOT.SelectedIndex = 0
            Me.cboRoundingTypeLate.SelectedIndex = 0
            Me.cboRoundingTypeEarlyOut.SelectedIndex = 0
            Me.cboRoundingTypeOT.SelectedIndex = 0
            Me.txtRoundingEarlyOT.Text = "00.00"
            Me.txtRoundingLate.Text = "00.00"
            Me.txtRoundingEarlyOut.Text = "00.00"
            Me.txtRoundingOT.Text = "00.00"
            Me.txtMaxValueEarlyOT.Text = "00.00"
            Me.txtMaxValueLate.Text = "00.00"
            Me.txtMaxValueEarlyOut.Text = "00.00"
            Me.txtMaxValueOT.Text = "00.00"
            Me.chkUseOvertimeOrder.Checked = False
            Me.chkOverrideCalculateEarlyOTTolerance.Checked = True
            Me.chkOverrideCalculateLateTolerance.Checked = True
            Me.chkOverrideCalculateEarlyOutTolerance.Checked = True
            Me.chkOverrideCalculateOTTolerance.Checked = True
            Me.dtpFirstTaxPeriod.Value = Me.ObjCekdata.GetFirstDateofTheMonth(Me.DtpAdmision.Value.Date)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.tblItemAttch.Clear()
            Me.tblItemTipeKontrak.Clear()
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim RowCountry As DseMasterKaryawan.tblmaster_kode_negaraRow
        Try
            With Me.ObjKaryawan
                Me.TxtPrimaryEmployee.Text = .Primary
                Me.TxtEmployeeCode.Text = .Code
                Me.TxtFullName.Text = .FullName
                Me.TxtPrefix.Text = .Prepix
                Me.TxtFirstName.Text = .FristName
                Me.TxtMiddleName.Text = .MiddleName
                Me.TxtLastName.Text = .LastName
                Me.TxtSuffix.Text = .Suffix
                Me.TxtNIK.Text = .NIK
                Me.TxtNotes.Text = .remark
                Me.cboStatus.SelectedIndex = .Aktif
                Me.DtpResignDate.Value = .ResignDate
                Me.TxtPrimaryDepartemen.Text = .PrimaryDep
                Me.TxtDepartemen.Text = Me.ObjDept.GetDeptCode(.PrimaryDep)
                Me.TxtPicturePath.Text = .PicturePath
                Me.DtpAdmision.Value = .AdmisionDate
                Me.dtpAppoinmentDate.Value = .AppointmentDate
                Me.TxtPrimaryPosition.Text = .PrimaryPosition
                Me.TxtPositionCode.Text = Me.ObjPost.GetCode(.PrimaryPosition)
                Me.TxtLegal_Astek_No.Text = .KPA_Astek_No
                Me.TxtLegal_KPJ_No.Text = .KPJ_no
                Me.TxtLegal_ID_Card_No.Text = .Id_Card_no
                Me.DtpLegal_Expired.Value = .Id_Card_Exp
                Me.TxtLegal_ID_Address.Text = .Id_Card_Add
                Me.TxtPersonal_NPWP.Text = .NPWP_No
                Me.CboLegal_Category.SelectedIndex = .NPWP_Category
                Me.TxtLegal_NPWP_Address.Text = .NPWP_Add
                Me.CboPersonal_MartialStatus.SelectedIndex = .MartialStatus
                Me.DtpPersonal_MTdate.Value = .MariageDAte
                Me.TxtPersonal_Education.Text = .Education
                Me.CboPersonal_Sex.SelectedIndex = .Sex
                Me.TxtLegal_Current_Address.Text = .CurentAdd
                Me.CboPersonal_Religion.SelectedIndex = .Religion
                Me.TxtPersonal_Phone.Text = .PhoneNo
                Me.CboEmployeeStatus.SelectedIndex = .EmployStatus
                Me.TxtPersonal_BirthPlace.Text = .BrithPalce
                Me.DtpPersonal_BPDate.Value = .DateBrith
                Me.CboPersonal_BloodType.SelectedIndex = .BloodType
                Me.TxtPersonal_Nationality.Text = .Nationality
                Me.TxtPersonal_Email.Text = .Email
                Me.TxtPersonal_Mother_Name.Text = .MotherName
                Me.TxtPersonal_Mother_Address.Text = .MotherAddres
                Me.TxtBank_Code.Text = .BankCode 'aul tambahan
                Me.TxtBank_Name.Text = .BankName
                Me.TxtBank_Branch.Text = .BankBranch
                Me.TxtBank_Acc_No.Text = .BankAccountNo
                Me.TxtBank_Acc_Holder.Text = .BankAccountHolder
                Me.TxtMedicalCenter.Text = .BalaiPengobatanJPK
                Me.ToolStripStatusLabel1.Text = "Created By: " + .CreatedBy
                Me.ToolStripStatusLabel2.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.ToolStripStatusLabel3.Text = "Modified By: " + .ModifiedBy
                Me.ToolStripStatusLabel4.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
                If .PicturePath <> "" Then
                    Dim fileExists As Boolean
                    fileExists = My.Computer.FileSystem.FileExists(.PicturePath)
                    If fileExists = False Then
                        Me.PictureBox1.Image = My.Resources.no_picture_fw
                        Me.AturPosisiKomponenGambar()
                        MsgBox("You don't have access to: " + .PicturePath + ". We can't show employee photo.", MsgBoxStyle.Information, "Information")
                    Else
                        Me.PictureBox1.Load(.PicturePath)
                        Me.AturPosisiKomponenGambar()
                    End If
                Else
                    Me.PictureBox1.Image = Nothing
                    Me.AturPosisiKomponenGambar()
                End If
                Me.TxtSalaryFormula.Text = .SalaryFormula
                If .IsNPWP = 0 Then
                    Me.ChkNoNPWP.Checked = False
                Else
                    Me.ChkNoNPWP.Checked = True
                End If
                If .ForeignEmployee = ClsMasterKaryawan.EnumForeign.No Then
                    Me.CboForeignEmployee.SelectedIndex = 0
                    Me.TxtPrimaryCountryCode.Text = 1
                    Me.CboCountryDesc.SelectedIndex = 0
                Else
                    Me.CboForeignEmployee.SelectedIndex = 1
                    Me.TxtPrimaryCountryCode.Text = .PrimaryCountry
                    If Me.TBLCountry.Rows.Count > 0 Then
                        For ij As Integer = 0 To Me.TBLCountry.Rows.Count - 1
                            RowCountry = Me.TBLCountry.Rows(ij)
                            If RowCountry.Primary = .PrimaryCountry Then
                                Me.CboCountryDesc.SelectedIndex = ij
                                Exit For
                            End If
                        Next
                    End If
                End If
                Me.TxtFingerprintID.Text = .FingerprintID
                If .OverrideClockIn = 0 Then
                    Me.ChkOverrideClockIn.Checked = False
                Else
                    Me.ChkOverrideClockIn.Checked = True
                End If
                If .DoesNotClockIn = 0 Then
                    Me.ChkDoesnNotClockIn.Checked = False
                Else
                    Me.ChkDoesnNotClockIn.Checked = True
                End If
                If .OverrideClockOut = 0 Then
                    Me.ChkOverrideClockOut.Checked = False
                Else
                    Me.ChkOverrideClockOut.Checked = True
                End If
                If .DoesNotClockOut = 0 Then
                    Me.ChkDoesnNotClockOut.Checked = False
                Else
                    Me.ChkDoesnNotClockOut.Checked = True
                End If

                If .OverrideEarlyOT = 0 Then
                    Me.ChkOverrideEarlyOT.Checked = False
                Else
                    Me.ChkOverrideEarlyOT.Checked = True
                End If
                If .IsCalculateEarlyOT = 0 Then
                    Me.CboIsCalculateEarlyOT.SelectedIndex = 1
                    Me.txtCalculateAfterEarlyOT.Text = Me.ObjCekdata.Convert_ToString(0)
                    Me.cboRoundingTypeEarlyOT.SelectedIndex = 0
                    Me.txtRoundingEarlyOT.Text = Me.ObjCekdata.Convert_ToString(0)
                    Me.txtMaxValueEarlyOT.Text = Me.ObjCekdata.Convert_ToString(0)
                Else
                    Me.CboIsCalculateEarlyOT.SelectedIndex = 0
                    Me.txtCalculateAfterEarlyOT.Text = Me.ObjCekdata.Convert_ToString(.CalculateAfterEarlyOT)
                    Me.cboRoundingTypeEarlyOT.SelectedIndex = .RoundingTypeEarlyOT
                    Me.txtRoundingEarlyOT.Text = Me.ObjCekdata.Convert_ToString(.RoundingValueEarlyOT)
                    Me.txtMaxValueEarlyOT.Text = Me.ObjCekdata.Convert_ToString(.MaxValueEarlyOT)
                End If

                If .OverrideLate = 0 Then
                    Me.ChkOverrideLate.Checked = False
                Else
                    Me.ChkOverrideLate.Checked = True
                End If
                If .IsCalculateLate = 0 Then
                    Me.CboIsCalculateLate.SelectedIndex = 1
                    Me.txtCalculateAfterLate.Text = Me.ObjCekdata.Convert_ToString(0)
                    Me.cboRoundingTypeLate.SelectedIndex = 0
                    Me.txtRoundingLate.Text = Me.ObjCekdata.Convert_ToString(0)
                    Me.txtMaxValueLate.Text = Me.ObjCekdata.Convert_ToString(0)
                Else
                    Me.CboIsCalculateLate.SelectedIndex = 0
                    Me.txtCalculateAfterLate.Text = Me.ObjCekdata.Convert_ToString(.CalculateAfterLate)
                    Me.cboRoundingTypeLate.SelectedIndex = .RoundingTypeLate
                    Me.txtRoundingLate.Text = Me.ObjCekdata.Convert_ToString(.RoundingValueLate)
                    Me.txtMaxValueLate.Text = Me.ObjCekdata.Convert_ToString(.MaxValueLate)
                End If

                If .OverrideEarlyOut = 0 Then
                    Me.ChkOverrideEarlyOut.Checked = False
                Else
                    Me.ChkOverrideEarlyOut.Checked = True
                End If
                If .IsCalculateEarlyOut = 0 Then
                    Me.CboIsCalculateEarlyOut.SelectedIndex = 1
                    Me.txtCalculateAfterEarlyOut.Text = Me.ObjCekdata.Convert_ToString(0)
                    Me.cboRoundingTypeEarlyOut.SelectedIndex = 0
                    Me.txtRoundingEarlyOut.Text = Me.ObjCekdata.Convert_ToString(0)
                    Me.txtMaxValueEarlyOut.Text = Me.ObjCekdata.Convert_ToString(0)
                Else
                    Me.CboIsCalculateEarlyOut.SelectedIndex = 0
                    Me.txtCalculateAfterEarlyOut.Text = Me.ObjCekdata.Convert_ToString(.CalculateAfterEarlyOut)
                    Me.cboRoundingTypeEarlyOut.SelectedIndex = .RoundingTypeEarlyOut
                    Me.txtRoundingEarlyOut.Text = Me.ObjCekdata.Convert_ToString(.RoundingValueEarlyOut)
                    Me.txtMaxValueEarlyOut.Text = Me.ObjCekdata.Convert_ToString(.MaxValueEarlyOut)
                End If

                If .OverrideOT = 0 Then
                    Me.ChkOverrideOT.Checked = False
                Else
                    Me.ChkOverrideOT.Checked = True
                End If
                If .IsCalculateOT = 0 Then
                    Me.CboIsCalculateOT.SelectedIndex = 1
                    Me.txtCalculateAfterOT.Text = Me.ObjCekdata.Convert_ToString(0)
                    Me.cboRoundingTypeOT.SelectedIndex = 0
                    Me.txtRoundingOT.Text = Me.ObjCekdata.Convert_ToString(0)
                    Me.txtMaxValueOT.Text = Me.ObjCekdata.Convert_ToString(0)
                Else
                    Me.CboIsCalculateOT.SelectedIndex = 0
                    Me.txtCalculateAfterOT.Text = Me.ObjCekdata.Convert_ToString(.CalculateAfterOT)
                    Me.cboRoundingTypeOT.SelectedIndex = .RoundingTypeOT
                    Me.txtRoundingOT.Text = Me.ObjCekdata.Convert_ToString(.RoundingValueOT)
                    Me.txtMaxValueOT.Text = Me.ObjCekdata.Convert_ToString(.MaxValueOT)
                End If

                If .OverrideCalculateEarlyOTTolerance = 1 Then
                    Me.chkOverrideCalculateEarlyOTTolerance.Checked = True
                Else
                    Me.chkOverrideCalculateEarlyOTTolerance.Checked = False
                End If
                If .OverrideCalculateLateTolerance = 1 Then
                    Me.chkOverrideCalculateLateTolerance.Checked = True
                Else
                    Me.chkOverrideCalculateLateTolerance.Checked = False
                End If
                If .OverrideCalculateEarlyOutTolerance = 1 Then
                    Me.chkOverrideCalculateEarlyOutTolerance.Checked = True
                Else
                    Me.chkOverrideCalculateEarlyOutTolerance.Checked = False
                End If
                If .OverrideCalculateOTTolerance = 1 Then
                    Me.chkOverrideCalculateOTTolerance.Checked = True
                Else
                    Me.chkOverrideCalculateOTTolerance.Checked = False
                End If

                Me.dtpFirstTaxPeriod.Value = .FirstTaxPeriod
                If Me.DGVKontrak.Rows.Count > 0 Then
                    Me.DGVKontrak.Rows(0).Selected = True
                End If

                If .UseOtDuration = 1 Then
                    Me.chkUseOvertimeOrder.Checked = True
                    Me.ChkOverrideOTOrder.Checked = True
                Else
                    Me.chkUseOvertimeOrder.Checked = False
                End If
                If .NotificationEmailOt = 1 Then
                    Me.chkSendNotifEmailOvertimeOrder.Checked = True
                Else
                    Me.chkSendNotifEmailOvertimeOrder.Checked = False
                End If

                If .OvverideOTOrder = 0 Then
                    If .UseOtDuration = 0 Then
                        Me.ChkOverrideOTOrder.Checked = False
                    End If
                Else
                    Me.ChkOverrideOTOrder.Checked = True
                End If

                If .LosesTaxObligation = 1 Then
                    Me.chkloses_tax_obligations.Checked = True
                Else
                    Me.chkloses_tax_obligations.Checked = False
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeObject()
        Try
            With Me.ObjKaryawan
                .Code = Me.TxtEmployeeCode.Text
                .FullName = Me.TxtFullName.Text
                .Prepix = Me.TxtPrefix.Text
                .FristName = Me.TxtFirstName.Text
                .MiddleName = Me.TxtMiddleName.Text
                .LastName = Me.TxtLastName.Text
                .Suffix = Me.TxtSuffix.Text
                .NIK = Me.TxtNIK.Text
                .remark = Me.TxtNotes.Text
                .Aktif = Me.cboStatus.SelectedIndex
                If Me.cboStatus.SelectedIndex = ClsMasterKaryawan.EnumAktif.NonActive Then
                    .ResignDate = Me.DtpResignDate.Value
                Else
                    .ResignDate = CDate("1/1/1950")
                End If
                .PrimaryDep = Me.TxtPrimaryDepartemen.Text
                .PicturePath = Me.TxtPicturePath.Text
                .AdmisionDate = Me.DtpAdmision.Value
                .AppointmentDate = Me.dtpAppoinmentDate.Value
                .PrimaryPosition = Me.TxtPrimaryPosition.Text
                .KPA_Astek_No = Me.TxtLegal_Astek_No.Text
                .KPJ_no = Me.TxtLegal_KPJ_No.Text
                .Id_Card_no = Me.TxtLegal_ID_Card_No.Text
                .Id_Card_Exp = Me.DtpLegal_Expired.Value
                .Id_Card_Add = Me.TxtLegal_ID_Address.Text
                .NPWP_No = Me.TxtPersonal_NPWP.Text
                .NPWP_Category = Me.CboLegal_Category.SelectedIndex
                .NPWP_Add = Me.TxtLegal_NPWP_Address.Text
                .MartialStatus = Me.CboPersonal_MartialStatus.SelectedIndex
                If Me.CboPersonal_MartialStatus.SelectedIndex = ClsMasterKaryawan.EnumMartialStatus.Married Then
                    .MariageDAte = Me.DtpPersonal_MTdate.Value
                Else
                    .MariageDAte = CDate("1/1/1950")
                End If
                .Education = Me.TxtPersonal_Education.Text
                .Sex = Me.CboPersonal_Sex.SelectedIndex
                .CurentAdd = Me.TxtLegal_Current_Address.Text
                .Religion = Me.CboPersonal_Religion.SelectedIndex
                .PhoneNo = Me.TxtPersonal_Phone.Text
                .EmployStatus = Me.CboEmployeeStatus.SelectedIndex
                .BrithPalce = Me.TxtPersonal_BirthPlace.Text
                .DateBrith = Me.DtpPersonal_BPDate.Value
                .BloodType = Me.CboPersonal_BloodType.SelectedIndex
                .Nationality = Me.TxtPersonal_Nationality.Text
                .Email = Me.TxtPersonal_Email.Text
                .MotherName = Me.TxtPersonal_Mother_Name.Text
                .MotherAddres = Me.TxtPersonal_Mother_Address.Text
                .BankCode = Me.TxtBank_Code.Text 'aul tambahan
                .BankName = Me.TxtBank_Name.Text
                .BankBranch = Me.TxtBank_Branch.Text
                .BankAccountNo = Me.TxtBank_Acc_No.Text
                .BankAccountHolder = Me.TxtBank_Acc_Holder.Text
                .BalaiPengobatanJPK = Me.TxtMedicalCenter.Text
                If Me.ChkNoNPWP.Checked = True Then
                    .IsNPWP = 1
                Else
                    .IsNPWP = 0
                End If
                If Me.CboForeignEmployee.SelectedIndex = 0 Then
                    .ForeignEmployee = ClsMasterKaryawan.EnumForeign.No
                    .PrimaryCountry = -1
                Else
                    .ForeignEmployee = ClsMasterKaryawan.EnumForeign.Yes
                    .PrimaryCountry = Me.TxtPrimaryCountryCode.Text
                End If
                .FingerprintID = Me.TxtFingerprintID.Text
                If Me.ChkOverrideClockIn.Checked = False Then
                    .OverrideClockIn = 0
                Else
                    .OverrideClockIn = 1
                End If
                If Me.ChkDoesnNotClockIn.Checked = False Then
                    .DoesNotClockIn = 0
                Else
                    .DoesNotClockIn = 1
                End If
                If Me.ChkOverrideClockOut.Checked = False Then
                    .OverrideClockOut = 0
                Else
                    .OverrideClockOut = 1
                End If
                If Me.ChkDoesnNotClockOut.Checked = False Then
                    .DoesNotClockOut = 0
                Else
                    .DoesNotClockOut = 1
                End If

                If Me.ChkOverrideEarlyOT.Checked = True Then
                    .OverrideEarlyOT = 1
                Else
                    .OverrideEarlyOT = 0
                End If
                If Me.CboIsCalculateEarlyOT.SelectedIndex = 0 Then
                    .IsCalculateEarlyOT = 1
                    .CalculateAfterEarlyOT = Me.ObjCekdata.Convert_ToMinutes(Me.txtCalculateAfterEarlyOT.Text)
                    .RoundingTypeEarlyOT = Me.cboRoundingTypeEarlyOT.SelectedIndex
                    .RoundingValueEarlyOT = Me.ObjCekdata.Convert_ToMinutes(Me.txtRoundingEarlyOT.Text)
                    .MaxValueEarlyOT = Me.ObjCekdata.Convert_ToMinutes(Me.txtMaxValueEarlyOT.Text)
                Else
                    .IsCalculateEarlyOT = 0
                    .CalculateAfterEarlyOT = 0
                    .RoundingTypeEarlyOT = 0
                    .RoundingValueEarlyOT = 0
                    .MaxValueEarlyOT = 0
                End If

                If Me.ChkOverrideLate.Checked = True Then
                    .OverrideLate = 1
                Else
                    .OverrideLate = 0
                End If
                If Me.CboIsCalculateLate.SelectedIndex = 0 Then
                    .IsCalculateLate = 1
                    .CalculateAfterLate = Me.ObjCekdata.Convert_ToMinutes(Me.txtCalculateAfterLate.Text)
                    .RoundingTypeLate = Me.cboRoundingTypeLate.SelectedIndex
                    .RoundingValueLate = Me.ObjCekdata.Convert_ToMinutes(Me.txtRoundingLate.Text)
                    .MaxValueLate = Me.ObjCekdata.Convert_ToMinutes(Me.txtMaxValueLate.Text)
                Else
                    .IsCalculateLate = 0
                    .CalculateAfterLate = 0
                    .RoundingTypeLate = 0
                    .RoundingValueLate = 0
                    .MaxValueLate = 0
                End If

                If Me.ChkOverrideEarlyOut.Checked = True Then
                    .OverrideEarlyOut = 1
                Else
                    .OverrideEarlyOut = 0
                End If
                If Me.CboIsCalculateEarlyOut.SelectedIndex = 0 Then
                    .IsCalculateEarlyOut = 1
                    .CalculateAfterEarlyOut = Me.ObjCekdata.Convert_ToMinutes(Me.txtCalculateAfterEarlyOut.Text)
                    .RoundingTypeEarlyOut = Me.cboRoundingTypeEarlyOut.SelectedIndex
                    .RoundingValueEarlyOut = Me.ObjCekdata.Convert_ToMinutes(Me.txtRoundingEarlyOut.Text)
                    .MaxValueEarlyOut = Me.ObjCekdata.Convert_ToMinutes(Me.txtMaxValueEarlyOut.Text)
                Else
                    .IsCalculateEarlyOut = 0
                    .CalculateAfterEarlyOut = 0
                    .RoundingTypeEarlyOut = 0
                    .RoundingValueEarlyOut = 0
                    .MaxValueEarlyOut = 0
                End If

                If Me.ChkOverrideOT.Checked = True Then
                    .OverrideOT = 1
                Else
                    .OverrideOT = 0
                End If
                If Me.CboIsCalculateOT.SelectedIndex = 0 Then
                    .IsCalculateOT = 1
                    .CalculateAfterOT = Me.ObjCekdata.Convert_ToMinutes(Me.txtCalculateAfterOT.Text)
                    .RoundingTypeOT = Me.cboRoundingTypeOT.SelectedIndex
                    .RoundingValueOT = Me.ObjCekdata.Convert_ToMinutes(Me.txtRoundingOT.Text)
                    .MaxValueOT = Me.ObjCekdata.Convert_ToMinutes(Me.txtMaxValueOT.Text)
                Else
                    .IsCalculateOT = 0
                    .CalculateAfterOT = 0
                    .RoundingTypeOT = 0
                    .RoundingValueOT = 0
                    .MaxValueOT = 0
                End If

                If Me.chkOverrideCalculateEarlyOTTolerance.Checked = True Then
                    .OverrideCalculateEarlyOTTolerance = 1
                Else
                    .OverrideCalculateEarlyOTTolerance = 0
                End If
                If Me.chkOverrideCalculateLateTolerance.Checked = True Then
                    .OverrideCalculateLateTolerance = 1
                Else
                    .OverrideCalculateLateTolerance = 0
                End If
                If Me.chkOverrideCalculateEarlyOutTolerance.Checked = True Then
                    .OverrideCalculateEarlyOutTolerance = 1
                Else
                    .OverrideCalculateEarlyOutTolerance = 0
                End If
                If Me.chkOverrideCalculateOTTolerance.Checked = True Then
                    .OverrideCalculateOTTolerance = 1
                Else
                    .OverrideCalculateOTTolerance = 0
                End If

                .FirstTaxPeriod = Me.dtpFirstTaxPeriod.Value
                If Me.chkUseOvertimeOrder.Checked = True Then
                    .UseOtDuration = 1
                Else
                    .UseOtDuration = 0
                End If
                If Me.chkSendNotifEmailOvertimeOrder.Checked = True Then
                    .NotificationEmailOt = 1
                Else
                    .NotificationEmailOt = 0
                End If
                If Me.ChkOverrideOTOrder.Checked = True Then
                    .OvverideOTOrder = 1
                Else
                    .OvverideOTOrder = 0
                End If
                If Me.chkloses_tax_obligations.Checked = True Then
                    .LosesTaxObligation = 1
                Else
                    .LosesTaxObligation = 0
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim ObjJurnal As New ClsTransaksiJournal1
        Dim ObjStockInOut As New clsTransaksiStockInOut1
        Dim RowDgv1 As DataGridViewRow
        Dim rowDgv2 As DataGridViewRow
        Dim rowDgv3 As DataGridViewRow
        Dim rowDgv4 As DataGridViewRow
        Dim rowDgv5 As DataGridViewRow
        Dim ObjCD As New ClsCekData
        Dim ApakahKaryawanAktif As Boolean = False
        Dim DummyDate As Date = #1/1/1950#
        Dim DummyDocNo As String = ""
        Dim DummyDocNo1 As String = ""
        Dim StartDate As Date = #1/1/1950#
        Dim EndDate As Date = #1/1/1950#

        Try
            'trim data
            Me.TxtFingerprintID.Text = Me.TxtFingerprintID.Text.Trim
            'tidak boleh kosong
            Me.TxtEmployeeCode.Text = Me.TxtEmployeeCode.Text.Trim
            KoleksiNamaData.Add("Employee Code")
            KoleksiTextBox.Add(Me.TxtEmployeeCode)
            KoleksiNamaData.Add("Name")
            KoleksiTextBox.Add(Me.TxtFullName)
            Me.TxtNIK.Text = Me.TxtNIK.Text.Trim
            KoleksiNamaData.Add("NIK")
            KoleksiTextBox.Add(Me.TxtNIK)
            KoleksiNamaData.Add("Department")
            KoleksiTextBox.Add(Me.TxtDepartemen)
            KoleksiNamaData.Add("Position")
            KoleksiTextBox.Add(Me.TxtPositionCode)
            If Me.ChkNoNPWP.Checked = True Then
                Me.TxtLegal_NPWP_Address.Text = Me.TxtLegal_NPWP_Address.Text.Trim
                KoleksiNamaData.Add("NPWP Address")
                KoleksiTextBox.Add(Me.TxtLegal_NPWP_Address)
            End If
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'tidak boleh double
            If Me.Baru = True Then
                If Me.ObjKaryawan.ApakahSudahAda(Me.TxtEmployeeCode.Text) = True Then
                    MsgBox("Duplication in Employee code, please change your code", MsgBoxStyle.Information, "Information")
                    Me.TxtEmployeeCode.Focus()
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.ObjKaryawan.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                If CodeLama <> Me.TxtEmployeeCode.Text Then
                    If Me.ObjKaryawan.ApakahSudahAda(Me.TxtEmployeeCode.Text) = True Then
                        MsgBox("Duplication in employee code, please change your code", MsgBoxStyle.Information, "Information")
                        Me.TxtEmployeeCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses(edit)
                Dim OldTimeStamp As Date = Me.ObjKaryawan.GetTimeStampByPrimary(Me.TxtPrimaryEmployee.Text)
                If Me.ObjKaryawan.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjKaryawan.CariDetail(Me.ObjKaryawan.Primary, Me, Me.dgvFamily, Me.DGVKompetensi)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'NIK tidak boleh sama
            If Me.ObjKaryawan.GetJmlKaryByNIK(Me.TxtPrimaryEmployee.Text, Me.TxtNIK.Text) > 0 Then
                MsgBox("NIK duplication, please change NIK.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'harus ada minimal 1 item kompetensi
            'If Me.ApakahPakaiCompetence = True Then
            '    If Me.DGVKompetensi.RowCount <= 1 Then
            '        MsgBox("Please input competence", MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If

            'cek data keluarga
            For Each RowDgv1 In Me.dgvFamily.Rows
                If RowDgv1.Cells(Me.ColName.Index).Value <> "" Then
                    If String.IsNullOrEmpty(RowDgv1.Cells(Me.ColHubKeluarga.Index).Value) = True Then
                        Me.TabControl2.SelectedIndex = 0
                        Me.dgvFamily.ClearSelection()
                        RowDgv1.Selected = True
                        MsgBox("Family relation can not be an empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    If String.IsNullOrEmpty(RowDgv1.Cells(Me.ColSex.Index).Value) = True Then
                        Me.TabControl2.SelectedIndex = 0
                        Me.dgvFamily.ClearSelection()
                        RowDgv1.Selected = True
                        MsgBox("Family member sex data can not be an empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    If String.IsNullOrEmpty(RowDgv1.Cells(Me.ColBirthPlace.Index).Value) = True Then
                        Me.TabControl2.SelectedIndex = 0
                        Me.dgvFamily.ClearSelection()
                        RowDgv1.Selected = True
                        MsgBox("Family member birth place data can not be an empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    If String.IsNullOrEmpty(RowDgv1.Cells(Me.ColBirthday.Index).Value) = True Or _
                    RowDgv1.Cells(Me.ColBirthday.Index).Value = "-" Then
                        Me.TabControl2.SelectedIndex = 0
                        Me.dgvFamily.ClearSelection()
                        RowDgv1.Selected = True
                        MsgBox("Family member date of birth data can not be an empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next

            'cek nilai pada dgv kompetensi
            For Each rowDgv2 In Me.DGVKompetensi.Rows
                If rowDgv2.Cells(Me.ColCompetence_Code.Index).Value <> "" Then
                    If String.IsNullOrEmpty(rowDgv2.Cells(Me.ColCompetence_Value.Index).Value) = True Then
                        Me.DGVKompetensi.ClearSelection()
                        rowDgv2.Selected = True
                        Me.TabControl2.SelectedIndex = 1
                        MsgBox("Value data for this employee  cannot empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next
            'cek data pada dgv kontrak
            For Each rowDgv3 In Me.DGVKontrak.Rows
                If rowDgv3.Cells(Me.ColContract_DocNo.Index).Value <> "" Then
                    If String.IsNullOrEmpty(rowDgv3.Cells(Me.ColContract_DocNo.Index).Value) = True Then
                        Me.DGVKontrak.ClearSelection()
                        rowDgv3.Selected = True
                        MsgBox("Work Contract document no. data for this employee cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                    If String.IsNullOrEmpty(rowDgv3.Cells(Me.ColContract_Date.Index).Value) = True Then
                        Me.DGVKontrak.ClearSelection()
                        rowDgv3.Selected = True
                        MsgBox("Work Contract date data for this employee cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If


                    If String.IsNullOrEmpty(rowDgv3.Cells(Me.ColContract_Start.Index).Value) = True Then
                        Me.DGVKontrak.ClearSelection()
                        rowDgv3.Selected = True
                        MsgBox("Work Contract Start date data for this employee cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If


                    If String.IsNullOrEmpty(rowDgv3.Cells(Me.ColContract_End.Index).Value) = True Then
                        Me.DGVKontrak.ClearSelection()
                        rowDgv3.Selected = True
                        MsgBox("Work Contract End date data for this employee cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next

            For Each rowItemAttch In tblItemAttch.Rows
                If rowItemAttch.desc = "" Or rowItemAttch.file_path = "" Then
                    For i As Integer = 0 To Me.DGVKontrak.Rows.Count - 1
                        If Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, i).Value = rowItemAttch.identifier_kontrak Then
                            Me.DGVKontrak.ClearSelection()
                            Me.DGVKontrak.Rows(i).Selected = True
                            If rowItemAttch.desc = "" Then
                                MsgBox("Attachment of Work Contract description data for this employee cannot be an empty string.", MsgBoxStyle.Information, "Information")
                            Else
                                MsgBox("Attachment of Work Contract file path data for this employee cannot be an empty string.", MsgBoxStyle.Information, "Information")
                            End If

                            For Each rowDgv4 In Me.DGVAttanchmentKontrak.Rows
                                If rowDgv4.Cells(Me.ColKontrakAtt_Description.Index).Value <> "" Then
                                    If String.IsNullOrEmpty(rowDgv4.Cells(Me.ColKontrakAtt_Description.Index).Value) = True Then
                                        Me.DGVAttanchmentKontrak.ClearSelection()
                                        Me.TabControl2.SelectTab(2)
                                        rowDgv4.Selected = True
                                        Return False
                                    End If

                                    If String.IsNullOrEmpty(rowDgv4.Cells(Me.ColKontrakAtt_LocationFile.Index).Value) = True Then
                                        Me.DGVAttanchmentKontrak.ClearSelection()
                                        Me.TabControl2.SelectTab(2)
                                        rowDgv4.Selected = True
                                        Return False
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
            Next
            For Each rowDgv5 In Me.DGVAttanchment.Rows
                If rowDgv5.Cells(Me.ColAtt_Description.Index).Value <> "" Then
                    If String.IsNullOrEmpty(rowDgv5.Cells(Me.ColAtt_Description.Index).Value) = True Then
                        Me.DGVAttanchment.ClearSelection()
                        MsgBox("Attachment description data for this employee cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Me.TabControl2.SelectTab(3)
                        rowDgv5.Selected = True
                        Return False
                    End If

                    If String.IsNullOrEmpty(rowDgv5.Cells(Me.ColAtt_LocationFile.Index).Value) = True Then
                        Me.DGVAttanchment.ClearSelection()
                        MsgBox("Attachment file path data for this employee cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Me.TabControl2.SelectTab(3)
                        rowDgv5.Selected = True
                        Return False
                    End If
                End If
            Next

            'admission date harus <= appointment date
            If Me.DtpAdmision.Value > Me.dtpAppoinmentDate.Value Then
                MsgBox("Admission date must be >= appointment date.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek batasan bila cbo rounding nya buka 'no rounding' maka rounding value nya tidak boleh 00.00
            If Me.cboRoundingTypeEarlyOT.SelectedIndex <> Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding Then
                If Me.txtRoundingEarlyOT.Text = "00.00" Then
                    MsgBox("Rounding type for early overtime : " & cboRoundingTypeEarlyOT.Text & _
                                           ", rounding value for early overtime cannot fill 00.00.", MsgBoxStyle.Information, "Information")
                    Me.txtRoundingEarlyOT.Focus()
                    Return False
                End If
            End If
            If Me.cboRoundingTypeLate.SelectedIndex <> Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding Then
                If Me.txtRoundingLate.Text = "00.00" Then
                    MsgBox("Rounding type for late : " & cboRoundingTypeLate.Text & _
                                          ", rounding value for late cannot fill 00.00.", MsgBoxStyle.Information, "Information")
                    Me.txtRoundingLate.Focus()
                    Return False
                End If
            End If
            If Me.cboRoundingTypeEarlyOut.SelectedIndex <> Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding Then
                If Me.txtRoundingEarlyOut.Text = "00.00" Then
                    MsgBox("Rounding type for early out : " & cboRoundingTypeEarlyOut.Text & _
                                          ", rounding value for early out cannot fill 00.00.", MsgBoxStyle.Information, "Information")
                    Me.txtRoundingEarlyOut.Focus()
                    Return False
                End If
            End If
            If Me.cboRoundingTypeOT.SelectedIndex <> Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding Then
                If Me.txtRoundingOT.Text = "00.00" Then
                    MsgBox("Rounding type for overtime : " & cboRoundingTypeOT.Text & _
                                         ", rounding value for overtime cannot fill 00.00.", MsgBoxStyle.Information, "Information")
                    Me.txtRoundingOT.Focus()
                    Return False
                End If
            End If

            If Me.ChkOverrideEarlyOT.Checked = True Then
                If ObjCD.Convert_ToMinutes(Me.txtCalculateAfterEarlyOT.Text) > 0 AndAlso _
               ObjCD.Convert_ToMinutes(Me.txtMaxValueEarlyOT.Text) > 0 AndAlso _
              ObjCD.Convert_ToMinutes(Me.txtCalculateAfterEarlyOT.Text) >= ObjCD.Convert_ToMinutes(Me.txtMaxValueEarlyOT.Text) Then
                    MsgBox("cannot fill calculate after early ot >= maximum value early ot", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If Me.ChkOverrideLate.Checked = True Then
                If ObjCD.Convert_ToMinutes(Me.txtCalculateAfterLate.Text) > 0 AndAlso _
               ObjCD.Convert_ToMinutes(Me.txtMaxValueLate.Text) > 0 AndAlso _
                  ObjCD.Convert_ToMinutes(Me.txtCalculateAfterLate.Text) >= ObjCD.Convert_ToMinutes(Me.txtMaxValueLate.Text) Then
                    MsgBox("cannot fill calculate after late >= maximum value late", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If Me.ChkOverrideEarlyOut.Checked = True Then
                If ObjCD.Convert_ToMinutes(Me.txtCalculateAfterEarlyOut.Text) > 0 AndAlso _
               ObjCD.Convert_ToMinutes(Me.txtMaxValueEarlyOut.Text) > 0 AndAlso _
              ObjCD.Convert_ToMinutes(Me.txtCalculateAfterEarlyOut.Text) >= ObjCD.Convert_ToMinutes(Me.txtMaxValueEarlyOut.Text) Then
                    MsgBox("cannot fill calculate after early out >= maximum value early out", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If Me.ChkOverrideOT.Checked = True Then
                If ObjCD.Convert_ToMinutes(Me.txtCalculateAfterOT.Text) > 0 AndAlso _
                   ObjCD.Convert_ToMinutes(Me.txtMaxValueOT.Text) > 0 AndAlso _
                   ObjCD.Convert_ToMinutes(Me.txtCalculateAfterOT.Text) >= ObjCD.Convert_ToMinutes(Me.txtMaxValueOT.Text) Then
                    MsgBox("cannot fill calculate after overtime >= maximum value overtime ", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If Me.cboStatus.SelectedIndex = 1 Then
                If Me.DtpAdmision.Value > Me.DtpResignDate.Value Then
                    MsgBox("Admision date value can't greater than resign date value", MsgBoxStyle.Information, "Information")
                    Me.DtpResignDate.Focus()
                    Return False
                End If
            End If

            'jika sudah ada perhitunga gaji, dengan tanggal end data lebih kecil dari dtp first tax period
            If Me.ObjKaryawan.GetFirstSalaryEndPeriod(Me.TxtPrimaryEmployee.Text) <> #1/1/1950# Then
                If Me.ObjCekdata.GetFirstDateofTheMonth(Me.dtpFirstTaxPeriod.Value).Date > Me.ObjKaryawan.GetFirstSalaryEndPeriod(Me.TxtPrimaryEmployee.Text) Then
                    MsgBox("First tax period period must be " + Format(Me.ObjKaryawan.GetFirstSalaryEndPeriod(Me.TxtPrimaryEmployee.Text), "MMM yyyy") + ".", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            If Me.ObjCekdata.GetFirstDateofTheMonth(Me.DtpAdmision.Value.Date) > Me.dtpFirstTaxPeriod.Value.Date Then
                MsgBox("First tax period period must be >= " + Format(Me.DtpAdmision.Value.Date, "MMM yyyy") + ".", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek apakah data lama resign
            If Me.Baru = False Then
                ApakahKaryawanAktif = Me.ObjKaryawan.ApakahAktif(Me.TxtPrimaryEmployee.Text)
                If ApakahKaryawanAktif = True And Me.cboStatus.SelectedIndex = 1 Then
                    'tadinya ga resign, diubah menjadi resign. tanggal resign minimal adalah tanggal hari terakhir gajian +1
                    DummyDate = Me.ObjKaryawan.GetLasSalaryEndPeriod(Me.TxtPrimaryEmployee.Text)
                    If Me.DtpResignDate.Value.Date <= DummyDate.Date Then
                        MsgBox("Minimum resign date is: " + Format(DummyDate + TimeSpan.FromDays(1), "dd MMM yyyy") + _
                        ", this employee has salary calculation before the respected date.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                ElseIf ApakahKaryawanAktif = False And Me.cboStatus.SelectedIndex = 0 Then
                    'tadinya tidak aktif, menjadi aktif
                    'pada tgl resign, tidak boleh ada transaksi gajian
                    DummyDate = Me.ObjKaryawan.GetResignDate(Me.TxtPrimaryEmployee.Text)
                    DummyDocNo = Me.ObjKaryawan.GetSalaryCalcNo(Me.TxtPrimaryEmployee.Text, DummyDate)
                    If DummyDocNo.Trim <> "" Then
                        MsgBox("This employee is considered resign on salary calculation: " + DummyDocNo + ", you can't set this employee status to 'Active'.")
                        Return False
                    End If
                End If

                If ApakahKaryawanAktif = False And Me.cboStatus.SelectedIndex = 1 Then
                    'tadi aktif, tetap tidak aktif, hanya tanggal resign yang di ubah, 
                    'maka tanggal resign tidak boleh pindah periode gajian
                    DummyDate = Me.ObjKaryawan.GetResignDate(Me.TxtPrimaryEmployee.Text)
                    If DummyDate.Date <> Me.DtpResignDate.Value.Date Then
                        DummyDocNo = Me.ObjKaryawan.GetSalaryCalcNo(Me.TxtPrimaryEmployee.Text, DummyDate)
                        DummyDocNo1 = Me.ObjKaryawan.GetSalaryCalcNo(Me.TxtPrimaryEmployee.Text, Me.DtpResignDate.Value.Date)

                        If DummyDocNo <> DummyDocNo1 Then
                            If DummyDocNo <> "" Then
                                StartDate = Me.ObjKaryawan.GetSalaryStartPeriod(Me.TxtPrimaryEmployee.Text, DummyDate)
                                EndDate = Me.ObjKaryawan.GetSalaryEndPeriod(Me.TxtPrimaryEmployee.Text, DummyDate)

                                MsgBox("This employee is considered resign on salary calculation: " + DummyDocNo + _
                                " (period: " + Format(StartDate, "dd MMM yyyy") + " - " + Format(EndDate, "dd MMM yyyy") + _
                                "), you can't change this employee resign date to a date outside this salary period range.")
                            Else
                                StartDate = Me.ObjKaryawan.GetSalaryStartPeriod(Me.TxtPrimaryEmployee.Text, Me.DtpResignDate.Value.Date)
                                EndDate = Me.ObjKaryawan.GetSalaryEndPeriod(Me.TxtPrimaryEmployee.Text, Me.DtpResignDate.Value.Date)
                                MsgBox("This employee refered by salary calculation: " + DummyDocNo1 + _
                                " (period: " + Format(StartDate, "dd MMM yyyy") + " - " + Format(EndDate, "dd MMM yyyy") + _
                                "), you can't change this employee resign date to this salary period range.")
                            End If

                            Return False
                        End If
                    End If
                End If
            End If
            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            If Me.TxtEmployeeCode.Enabled = True Then
                Dim ijk As New Core_Module.FrmMasterKaryawan_DetailNama
                ijk.frm = Me
                ijk.TxtFirstName.Text = Me.TxtFirstName.Text
                ijk.TxtLastName.Text = Me.TxtLastName.Text
                ijk.TxtMiddleName.Text = Me.TxtMiddleName.Text
                ijk.TxtPrefix.Text = Me.TxtPrefix.Text
                ijk.TxtSuffix.Text = Me.TxtSuffix.Text
                ijk.CancelButton = ijk.btnCancel
                ijk.BackColorAll = Me.BackColorAll
                ijk.ShowDialog()
                ijk.Dispose()
                Me.SetName()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetName()
        If Me.TxtPrefix.Text = "" Then
            Me.TxtFullName.Text = Me.TxtFirstName.Text + " " + Me.TxtMiddleName.Text _
            + " " + Me.TxtLastName.Text + " " + Me.TxtSuffix.Text
        Else
            Me.TxtFullName.Text = Me.TxtPrefix.Text + " " + Me.TxtFirstName.Text + " " + Me.TxtMiddleName.Text _
            + " " + Me.TxtLastName.Text + " " + Me.TxtSuffix.Text
        End If
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            If Me.TxtEmployeeCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtEmployeeCode.Text <> "" Then
                Dim i As New FrmMasterKaryawan_Cetak1
                i.mPrimaryKaryawan = Me.TxtPrimaryEmployee.Text
                i.ShowDialog()
                i.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterKaryawan_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtEmployeeCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObject()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjKaryawan.SimpanBaru(Me.UserCode, Me.dgvFamily, Me.DGVKompetensi, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjKaryawan.SimpanPerubahan(Me.UserCode, Me.dgvFamily, Me.DGVKompetensi, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
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

    Private Sub FrmMasterKaryawan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.MinimizeBox = True
            Me.NonAktifkanDetail()
            Me.KosongkanField()
            Me.DGVKompetensi.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvFamily.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVKontrak.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVAttanchmentKontrak.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVAttanchment.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If Me.ApakahPakaiCompetence = False Then
                Me.TabControl2.TabPages.Remove(Me.TabPage6)
            End If
            If Me.ChkOverrideClockIn.Checked = True Then
                Me.ChkOverrideClockIn.Checked = False
                Me.ChkOverrideClockIn.Checked = True
            ElseIf Me.ChkOverrideClockIn.Checked = False Then
                Me.ChkOverrideClockIn.Checked = True
                Me.ChkOverrideClockIn.Checked = False
            End If
            If Me.ChkOverrideClockOut.Checked = True Then
                Me.ChkOverrideClockOut.Checked = False
                Me.ChkOverrideClockOut.Checked = True
            ElseIf Me.ChkOverrideClockOut.Checked = False Then
                Me.ChkOverrideClockOut.Checked = True
                Me.ChkOverrideClockOut.Checked = False
            End If
            ObjKaryawanType.isiCboEmployeeTypeContrack(Me.TBLKontrak)
            RefreshCbo()
            Me.dtpFirstTaxPeriod.Value = Me.ObjCekdata.GetFirstDateofTheMonth(Me.dtpFirstTaxPeriod.Value)

            Me.chkUseOvertimeOrder.Checked = False
            Me.chkSendNotifEmailOvertimeOrder.Checked = False

            Me.Label29.Text = "Health BPJS No"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub RefreshCbo()
        Try
            If Me.ChkOverrideClockIn.Checked = True Then
                Me.ChkOverrideClockIn.Checked = False
                Me.ChkOverrideClockIn.Checked = True
            ElseIf Me.ChkOverrideClockIn.Checked = False Then
                Me.ChkOverrideClockIn.Checked = True
                Me.ChkOverrideClockIn.Checked = False
            End If
            If Me.ChkOverrideClockOut.Checked = True Then
                Me.ChkOverrideClockOut.Checked = False
                Me.ChkOverrideClockOut.Checked = True
            ElseIf Me.ChkOverrideClockOut.Checked = False Then
                Me.ChkOverrideClockOut.Checked = True
                Me.ChkOverrideClockOut.Checked = False
            End If

            If Me.ChkOverrideEarlyOT.Checked = True Then
                Me.ChkOverrideEarlyOT.Checked = False
                Me.ChkOverrideEarlyOT.Checked = True
            ElseIf Me.ChkOverrideEarlyOT.Checked = False Then
                Me.ChkOverrideEarlyOT.Checked = True
                Me.ChkOverrideEarlyOT.Checked = False
            End If
            If Me.ChkOverrideLate.Checked = True Then
                Me.ChkOverrideLate.Checked = False
                Me.ChkOverrideLate.Checked = True
            ElseIf Me.ChkOverrideLate.Checked = False Then
                Me.ChkOverrideLate.Checked = True
                Me.ChkOverrideLate.Checked = False
            End If
            If Me.ChkOverrideEarlyOut.Checked = True Then
                Me.ChkOverrideEarlyOut.Checked = False
                Me.ChkOverrideEarlyOut.Checked = True
            ElseIf Me.ChkOverrideEarlyOut.Checked = False Then
                Me.ChkOverrideEarlyOut.Checked = True
                Me.ChkOverrideEarlyOut.Checked = False
            End If
            If Me.ChkOverrideOT.Checked = True Then
                Me.ChkOverrideOT.Checked = False
                Me.ChkOverrideOT.Checked = True
            ElseIf Me.ChkOverrideOT.Checked = False Then
                Me.ChkOverrideOT.Checked = True
                Me.ChkOverrideOT.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.Baru = True
            Me.KosongkanField()
            Me.Aktifkandetail()
            Me.TxtEmployeeCode.Focus()
            Me.DtpPersonal_BPDate.Value = Now
            Me.DtpPersonal_MTdate.Value = Now
            Me.TabControl1.SelectedIndex = 0
            Me.TabControl2.SelectedIndex = 0
            Me.ToolStripStatusLabel1.Text = "Created By: " + Me.UserCode
            Me.ToolStripStatusLabel2.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.ToolStripStatusLabel3.Text = "Modified By: " + Me.UserCode
            Me.ToolStripStatusLabel4.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            'cek data kosong
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObject()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    If Me.ObjKaryawan.ApakahSesuaiDenganJobList(Me.TxtPrimaryPosition.Text, Me) = False Then
                        If Me.ApakahPakaiCompetence = True Then
                            hasil = MsgBox("Job position in " + Me.ObjPost.GetDesc(Me.TxtPrimaryPosition.Text) + " requires following competences : " + vbNewLine + _
                            Me.ObjKaryawan.GetListKompeten(Me.TxtPrimaryPosition.Text) + "Continues to save data ?", MsgBoxStyle.YesNo, "Confirm")
                        Else
                            hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
                        End If
                    Else
                        hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
                    End If
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjKaryawan.SimpanBaru(Me.UserCode, Me.dgvFamily, Me.DGVKompetensi, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    If Me.ObjKaryawan.ApakahSesuaiDenganJobList(Me.TxtPrimaryPosition.Text, Me) = False Then
                        If Me.ApakahPakaiCompetence = True Then
                            hasil = MsgBox("Job position in " + Me.ObjPost.GetDesc(Me.TxtPrimaryPosition.Text) + " requires following competences : " + vbNewLine + _
                            Me.ObjKaryawan.GetListKompeten(Me.TxtPrimaryPosition.Text) + "Continues to save data ?", MsgBoxStyle.YesNo, "Confirm")
                        Else
                            hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
                        End If
                    Else
                        hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
                    End If
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjKaryawan.SimpanPerubahan(Me.UserCode, Me.dgvFamily, Me.DGVKompetensi, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AturPosisiKomponenGambar()
        Try
            If Me.TxtPicturePath.Text <> "" Then
                Me.PictureBox1.BorderStyle = BorderStyle.None
                Me.PictureBox1.BackColor = Color.Transparent
            Else
                Me.PictureBox1.BorderStyle = BorderStyle.FixedSingle
                Me.PictureBox1.BackColor = Color.White
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        Try
            If Me.TxtNotes.Enabled = True Then
                If Me.OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    'load data gambar ke picture box
                    Me.PictureBox1.Load(Me.OpenFileDialog1.FileName)
                    Me.TxtPicturePath.Text = Me.OpenFileDialog1.FileName
                    'atur posisi komponen untuk gambar
                    Me.AturPosisiKomponenGambar()
                Else
                    If Me.TxtPicturePath.Text <> "" Then
                        'load data gambar ke picture box
                        Me.PictureBox1.Load(Me.TxtPicturePath.Text)
                        Me.AturPosisiKomponenGambar()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        Try
            If Me.TxtNotes.Enabled = True Then
                Me.TxtPicturePath.Text = ""
                Me.PictureBox1.Image = Nothing
                Me.AturPosisiKomponenGambar()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtEmployeeCode.Text <> "" And Me.TxtEmployeeCode.Enabled = True Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.NonAktifkanDetail()
                        Me.KosongkanField()
                    Else
                        Me.ObjKaryawan.CariDetail(Me.TxtPrimaryEmployee.Text, Me, Me.dgvFamily, Me.DGVKompetensi)
                        Me.SalinKeField()
                        Me.TxtPersonal_NPWP.Enabled = False
                        Me.TxtLegal_NPWP_Address.Enabled = False
                        Me.NonAktifkanDetail()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
                Me.NonAktifkanDetail()
            End If
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.TxtEmployeeCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtEmployeeCode.Text <> "" Then
                Me.Baru = False
                Me.ObjKaryawan.CariDetail(Me.TxtPrimaryEmployee.Text, Me, Me.dgvFamily, Me.DGVKompetensi)
                Me.Aktifkandetail()
                Me.TxtEmployeeCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.TxtEmployeeCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtEmployeeCode.Text <> "" Then
                If Me.ObjKaryawan.ApakahBolehHapus(Me.TxtPrimaryEmployee.Text, Me.UserCode) = True Then
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjKaryawan.Hapus(Me.TxtPrimaryEmployee.Text, Me.UserCode) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmMasterKaryawanCari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryEmploye = -1 Then
                If Me.ObjKaryawan.CariDetail(hj.PrimaryEmploye, Me, Me.dgvFamily, Me.DGVKompetensi) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                    Me.NonAktifkanDetail()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnFindDepartemen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindDepartemen.Click
        Dim FrmCariDept As New FrmMasterDepartementCari
        Try
            FrmCariDept.BackColorAll = Me.BackColorAll
            If Me.TxtPrimaryDepartemen.Text = -1 And Me.TxtDepartemen.Text = "" Then
                FrmCariDept.ShowDialog()
            Else
                FrmCariDept.txtCari.Text = Me.TxtDepartemen.Text
                FrmCariDept.ShowDialog()
            End If
            If CDbl(FrmCariDept.PrimaryDepartement) > 0 Then
                Me.TxtPrimaryDepartemen.Text = FrmCariDept.PrimaryDepartement
                Me.TxtDepartemen.Text = Me.ObjDept.GetDeptCode(FrmCariDept.PrimaryDepartement)
                Me.TxtPositionCode.Focus()
            Else
                Me.TxtPrimaryDepartemen.Text = -1
                Me.TxtDepartemen.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCariDept.Dispose()
        End Try
    End Sub

    Private Sub TxtDepartemen_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDepartemen.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjDept.ApakahSudahAda(Me.TxtDepartemen.Text) = True Then
                    Me.TxtPrimaryDepartemen.Text = Me.ObjDept.GetPrimary(Me.TxtDepartemen.Text)
                    Me.TxtPositionCode.Focus()
                Else
                    Me.btnFindDepartemen.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDepartemen_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDepartemen.Leave
        Try
            If Me.TxtDepartemen.Text <> "" Then
                If Me.ObjDept.ApakahSudahAda(Me.TxtDepartemen.Text) = True Then
                    Me.TxtPrimaryDepartemen.Text = Me.ObjDept.GetPrimary(Me.TxtDepartemen.Text)
                Else
                    Me.btnFindDepartemen.PerformClick()
                End If
            Else
                Me.TxtPrimaryDepartemen.Text = -1
                Me.TxtDepartemen.Text = ""
            End If
            Me.TxtDepartemen.BackColor = Me.TxtPrimaryDepartemen.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindPosition_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindPosition.Click
        Dim FrmCariPos As New FrmMasterJobPosition_Cari
        Try
            Me.buffer = Me.TxtPositionCode.Text
            FrmCariPos.BackColorAll = Me.BackColorAll
            If Me.TxtPrimaryPosition.Text = -1 And Me.TxtPositionCode.Text = "" Then
                FrmCariPos.ShowDialog()
            Else
                FrmCariPos.txtCari.Text = Me.TxtPositionCode.Text
                FrmCariPos.ShowDialog()
            End If
            If CDbl(FrmCariPos.mPrimary) > 0 Then
                Me.TxtPrimaryPosition.Text = FrmCariPos.mPrimary
                Me.TxtPositionCode.Text = Me.ObjPost.GetCode(FrmCariPos.mPrimary)
                Me.DGVKompetensi.AllowUserToAddRows = True
                Me.DGVKompetensi.ReadOnly = False
                If Me.DGVKompetensi.Rows.Count <= 1 Then
                    If Me.TxtPositionCode.Text <> Me.buffer Then
                        Me.LoadKompetensi(Me.TxtPrimaryPosition.Text, Me.ApakahPakaiCompetence)
                    Else
                        If Me.DGVKompetensi.Rows.Count <= 1 Then
                            Me.LoadKompetensi(Me.TxtPrimaryPosition.Text, Me.ApakahPakaiCompetence)
                        End If
                    End If
                End If
                Me.CboEmployeeStatus.Focus()
            Else
                Me.TxtPrimaryPosition.Text = -1
                Me.TxtPositionCode.Text = ""
                Me.DGVKompetensi.AllowUserToDeleteRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCariPos.Dispose()
        End Try
    End Sub

    Private Sub TxtPositionCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPositionCode.KeyDown
        Try
            Me.buffer = Me.TxtPositionCode.Text
            If e.KeyValue = 13 Then
                If Me.ObjPost.ApakahSudahAda(Me.TxtPositionCode.Text) = True Then
                    Me.TxtPrimaryPosition.Text = Me.ObjPost.GetPrimary(Me.TxtPositionCode.Text)
                    Me.CboEmployeeStatus.Focus()
                Else
                    Me.BtnFindPosition.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPositionCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPositionCode.Leave
        Dim FrmCariPos As New FrmMasterJobPosition_Cari
        Try
            Me.buffer = Me.TxtPositionCode.Text
            If Me.TxtPositionCode.Text <> "" Then
                If Me.ObjPost.ApakahSudahAda(Me.TxtPositionCode.Text) = True Then
                    Me.TxtPrimaryPosition.Text = Me.ObjPost.GetPrimary(Me.TxtPositionCode.Text)
                    Me.DGVKompetensi.AllowUserToAddRows = True
                    Me.DGVKompetensi.AllowUserToDeleteRows = True
                    If Me.DGVKompetensi.Rows.Count <= 1 Then
                        If Me.TxtPositionCode.Text <> Me.buffer Then
                            Me.LoadKompetensi(Me.TxtPrimaryPosition.Text, Me.ApakahPakaiCompetence)
                        Else
                            If Me.DGVKompetensi.Rows.Count <= 1 Then
                                Me.LoadKompetensi(Me.TxtPrimaryPosition.Text, Me.ApakahPakaiCompetence)
                            End If
                        End If
                    End If
                Else
                    Me.BtnFindPosition.PerformClick()
                End If
            Else
                Me.TxtPrimaryPosition.Text = -1
                Me.TxtPositionCode.Text = ""
                'Me.DGVKompetensi.AllowUserToAddRows = False
                Me.DGVKompetensi.AllowUserToDeleteRows = True
                If Me.DGVKompetensi.Rows.Count > 1 Then
                    If MsgBox("Clear competence data?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                        Me.DGVKompetensi.Rows.Clear()
                    End If
                End If
            End If
            Me.TxtPositionCode.BackColor = Me.TxtPrimaryPosition.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboPersonal_MartialStatus_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboPersonal_MartialStatus.SelectedIndexChanged
        Try
            If Me.CboPersonal_MartialStatus.SelectedIndex = ClsMasterKaryawan.EnumMartialStatus.Married Then
                Me.DtpPersonal_MTdate.Visible = True
                Me.LblDateOfMarriage.Visible = False
            ElseIf Me.CboPersonal_MartialStatus.SelectedIndex = ClsMasterKaryawan.EnumMartialStatus._Single Then
                Me.DtpPersonal_MTdate.Visible = False
                Me.LblDateOfMarriage.Visible = True
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub dgvFamily_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvFamily.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.dgvFamily(Me.ColName.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColHubKeluarga.Index Or e.ColumnIndex = Me.ColIsDependent.Index Or _
                e.ColumnIndex = Me.ColEducation.Index Or e.ColumnIndex = Me.ColContactNo.Index Or _
                e.ColumnIndex = Me.ColBirthPlace.Index Or e.ColumnIndex = Me.ColBirthday.Index Or _
                e.ColumnIndex = Me.ColRemarks.Index Or e.ColumnIndex = Me.ColName.Index Or e.ColumnIndex = Me.ColSex.Index Or _
                e.ColumnIndex = Me.ColGolDarah.Index Then
                    Me.buffer = Me.dgvFamily.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColName.Index Then
                    Me.buffer = Me.dgvFamily.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvFamily_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvFamily.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvFamily.Rows(e.RowIndex)
        Try
            If IsNothing(Me.dgvFamily.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If Me.buffer = Me.dgvFamily.Item(e.ColumnIndex, e.RowIndex).Value Then
                    Exit Sub
                End If
                With row
                    'nama
                    If e.ColumnIndex = Me.ColName.Index Then
                        If String.IsNullOrEmpty(Me.dgvFamily.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            Me.mDgv = Me.dgvFamily
                            Me.mEraseRowIndex = row.Index
                            Me.BgwDeleteRowDgv.RunWorkerAsync()
                        Else
                            Me.dgvFamily.Item(Me.ColIdentifier.Index, e.RowIndex).Value = Me.ObjItem.GetIdentifierUrutan(Me.dgvFamily, Me.ColIdentifier.Index, _
                            Me.Baru, Me.TxtPrimaryEmployee.Text, e.RowIndex)
                        End If
                    End If
                    'DOB
                    If e.ColumnIndex = Me.ColBirthday.Index Then
                        'set region ke indonesia
                        My.Application.ChangeUICulture("id-ID")
                        My.Application.ChangeCulture("id-ID")
                        ''Thread.CurrentThread.CurrentCulture = New CultureInfo("id-ID", True)
                        If IsDate(Me.dgvFamily.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            If String.IsNullOrEmpty(Me.dgvFamily.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                Me.dgvFamily.Item(e.ColumnIndex, e.RowIndex).Value = "-"
                            Else
                                MsgBox("Must input date format.", MsgBoxStyle.Information, "Information")
                                Me.dgvFamily.Item(e.ColumnIndex, e.RowIndex).Value = "-"
                            End If
                        Else
                            Me.dgvFamily.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.dgvFamily.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                        End If
                        'kembalikan ke region defaultnya
                        My.Application.ChangeCulture("en-US")
                        My.Application.ChangeUICulture("en-US")
                        ''Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", True)
                    End If
                End With
            Else
               
                Me.dgvFamily.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                If String.IsNullOrEmpty(Me.buffer) = True Then
                    Me.mDgv = Me.dgvFamily
                    Me.mEraseRowIndex = e.RowIndex
                    Me.BgwDeleteRowDgv.RunWorkerAsync()
                Else
                    Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
                ObjCekdata.SetNumber(Me.dgvFamily)
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvFamily_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvFamily.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            Me.dgvFamily.TabStop = True
            rowDGV = Me.dgvFamily.Rows(Me.dgvFamily.RowCount - 1)
            With rowDGV
                .Cells(Me.ColName.Index).Value = ""
            End With
            Dim i As Integer = 1
            Dim CboDgv As DataGridViewComboBoxCell
            For Each rowDGV In dgvFamily.Rows
                rowDGV.Cells(Me.ColNo.Index).Value = i
                If String.IsNullOrEmpty(rowDGV.Cells(Me.ColName.Index).Value) = True Then
                    CboDgv = rowDGV.Cells(Me.ColHubKeluarga.Index)
                    rowDGV.Cells(Me.ColHubKeluarga.Index).Value = CboDgv.Items.Item(0)
                    rowDGV.Cells(Me.ColIsDependent.Index).Value = False '/Penambahan value terkait error saat pindah kolom malah menghapus baris
                    CboDgv = rowDGV.Cells(Me.ColSex.Index)
                    rowDGV.Cells(Me.ColSex.Index).Value = CboDgv.Items.Item(0)
                    CboDgv = rowDGV.Cells(Me.ColGolDarah.Index)
                    rowDGV.Cells(Me.ColGolDarah.Index).Value = CboDgv.Items.Item(0)
                    rowDGV.Cells(Me.ColBirthday.Index).Value = "-"
                End If
                i = i + 1
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DGVKompetensi_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVKompetensi.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.DGVKompetensi(Me.ColCompetence_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColCompetence_Value.Index Or e.ColumnIndex = Me.ColCompetence_NextCertificationDate.Index Or _
                e.ColumnIndex = Me.ColCompetence_Code.Index Or e.ColumnIndex = Me.ColCompetence_Remarks.Index Then
                    Me.buffer = Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value
                    If e.ColumnIndex = Me.ColCompetence_NextCertificationDate.Index Then
                        If Me.ObjPost.GetPeriodic(Me.TxtPrimaryPosition.Text, Me.DGVKompetensi.Item(Me.ColCompetence_Primary.Index, e.RowIndex).Value) = 0 Then
                            e.Cancel = True
                        End If
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColCompetence_Code.Index Then
                    Me.buffer = Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVKompetensi_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVKompetensi.CellEndEdit
        Dim row As DataGridViewRow = Me.DGVKompetensi.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim Primary As Double
        Dim DefPrec As Integer = 0
        Dim DataType As Integer
        Try
            With row
                '===code
                If IsNothing(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.buffer = Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColCompetence_Code.Index Then
                        If String.IsNullOrEmpty(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            If Me.ObjKompetensi.ApakahSudahAda(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                Primary = Me.ObjKompetensi.GetPrimary(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value)
                                index = Me.GetIndexItemInput(Primary)
                                If index < 0 Then
                                    DefPrec = Me.ObjKompetensi.GetDefPrec(Primary)
                                    DataType = Me.ObjKompetensi.GetDataType(Primary)
                                    Me.DGVKompetensi.Item(Me.ColCompetence_Primary.Index, e.RowIndex).Value = Primary
                                    Me.DGVKompetensi.Item(Me.ColCompetence_Code.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDataCode(Primary)
                                    Me.DGVKompetensi.Item(Me.ColCempetence_Description.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDesc(Primary)
                                    If Me.ObjPost.GetJmlItem(Me.TxtPrimaryPosition.Text, Primary) > 0 Then
                                        Me.DGVKompetensi.Item(Me.ColCompetence_StandarValue.Index, e.RowIndex).Value = _
                                        Me.ObjPost.GetValue_CustomFormatted(Me.TxtPrimaryPosition.Text, Primary, Me.ObjCekdata, DataType)
                                    Else
                                        Me.DGVKompetensi.Item(Me.ColCompetence_StandarValue.Index, e.RowIndex).Value = _
                                        Me.ObjKompetensi.GetValue_CustomFormatted(Primary, Me.ObjCekdata, DataType)
                                    End If
                                    Me.DGVKompetensi.Item(Me.ColCempetence_DataType.Index, e.RowIndex).Value = DataType
                                    If Me.ObjPost.GetPeriodic(Me.TxtPrimaryPosition.Text, Primary) = 1 Then
                                        Me.DGVKompetensi.Item(Me.ColCompetence_NextCertificationDate.Index, e.RowIndex).Value = Format(Now, "dd MMM yyyy")
                                    Else
                                        Me.DGVKompetensi.Item(Me.ColCompetence_NextCertificationDate.Index, e.RowIndex).Value = "-"
                                    End If
                                    Me.DGVKompetensi.Item(Me.ColCompetence_Value.Index, e.RowIndex).Value = ""
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                                        Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        If Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                            Me.mDgv = Me.DGVKompetensi
                                            Me.mEraseRowIndex = row.Index
                                            Me.BgwDeleteRowDgv.RunWorkerAsync()
                                        End If
                                        Exit Sub
                                    End If
                                End If
                            Else
                                Dim ij As New FrmMasterKompetensi_Cari
                                ij.txtCari.Text = Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value
                                ij.ShowDialog()
                                index = Me.GetIndexItemInput(ij.mPrimary)
                                If ij.mPrimary = -1 Then
                                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                    Me.mDgv = Me.DGVKompetensi
                                    Me.mEraseRowIndex = row.Index
                                    Me.BgwDeleteRowDgv.RunWorkerAsync()
                                Else
                                    If index < 0 Then
                                        DefPrec = Me.ObjKompetensi.GetDefPrec(ij.mPrimary)
                                        DataType = Me.ObjKompetensi.GetDataType(ij.mPrimary)
                                        Me.DGVKompetensi.Item(Me.ColCompetence_Primary.Index, e.RowIndex).Value = ij.mPrimary
                                        Me.DGVKompetensi.Item(Me.ColCompetence_Code.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDataCode(ij.mPrimary)
                                        Me.DGVKompetensi.Item(Me.ColCempetence_Description.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDesc(ij.mPrimary)
                                        If Me.ObjPost.GetJmlItem(Me.TxtPrimaryPosition.Text, ij.mPrimary) > 0 Then
                                            Me.DGVKompetensi.Item(Me.ColCompetence_StandarValue.Index, e.RowIndex).Value = _
                                            Me.ObjPost.GetValue_CustomFormatted(Me.TxtPrimaryPosition.Text, ij.mPrimary, Me.ObjCekdata, DataType)
                                        Else
                                            Me.DGVKompetensi.Item(Me.ColCompetence_StandarValue.Index, e.RowIndex).Value = _
                                            Me.ObjKompetensi.GetValue_CustomFormatted(ij.mPrimary, Me.ObjCekdata, DataType)
                                        End If
                                        Me.DGVKompetensi.Item(Me.ColCempetence_DataType.Index, e.RowIndex).Value = DataType
                                        If Me.ObjPost.GetPeriodic(Me.TxtPrimaryPosition.Text, ij.mPrimary) = 1 Then
                                            Me.DGVKompetensi.Item(Me.ColCompetence_NextCertificationDate.Index, e.RowIndex).Value = Format(Now, "dd MMM yyyy")
                                        Else
                                            Me.DGVKompetensi.Item(Me.ColCompetence_NextCertificationDate.Index, e.RowIndex).Value = "-"
                                        End If
                                        Me.DGVKompetensi.Item(Me.ColCompetence_Value.Index, e.RowIndex).Value = ""
                                    Else
                                        If e.RowIndex <> index Then
                                            MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                                            Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                            If Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                                Me.mDgv = Me.DGVKompetensi
                                                Me.mEraseRowIndex = row.Index
                                                Me.BgwDeleteRowDgv.RunWorkerAsync()
                                            End If
                                            Exit Sub
                                        Else
                                            Me.DGVKompetensi.Item(Me.ColCompetence_Code.Index, e.RowIndex).Value = Me.ObjKompetensi.GetDataCode(ij.mPrimary)
                                        End If
                                    End If
                                End If
                                ij = Nothing
                            End If
                        Else
                            Me.mDgv = Me.DGVKompetensi
                            Me.mEraseRowIndex = row.Index
                            Me.BgwDeleteRowDgv.RunWorkerAsync()
                        End If
                    End If
                    '===value
                    If e.ColumnIndex = Me.ColCompetence_Value.Index Then
                        If String.IsNullOrEmpty(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            DefPrec = Me.ObjKompetensi.GetDefPrec(Me.DGVKompetensi.Item(Me.ColCompetence_Primary.Index, e.RowIndex).Value)
                            DataType = Me.ObjKompetensi.GetDataType(Me.DGVKompetensi.Item(Me.ColCompetence_Primary.Index, e.RowIndex).Value)
                            If DataType = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                                If Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "A" Or _
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "B" Or _
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "C" Or _
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "D" Or _
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "E" Or _
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "F" Then
                                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = _
                                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                                Else
                                    MsgBox("Must input A , B , C ,D , E ,F", MsgBoxStyle.Information, "Information")
                                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            ElseIf DataType = ClsMasterKompetensi.EnumDataType.Number Then
                                If Me.ObjCekdata.CekNumerik(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                    MsgBox("Must input number format", MsgBoxStyle.Information, "Information")
                                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Else
                                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = _
                                    Me.ObjKompetensi.SetFormatDataCustom(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value, DataType, DefPrec, Me.ObjCekdata)
                                End If
                            ElseIf DataType = ClsMasterKompetensi.EnumDataType.Tanggal Then
                                'set region ke indonesia
                                If IsDate(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                    MsgBox("Must input date format.", MsgBoxStyle.Information, "Information")
                                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Else
                                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = _
                                    Me.ObjKompetensi.SetFormatDataCustom(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value, DataType, DefPrec, Me.ObjCekdata)
                                End If
                            ElseIf DataType = ClsMasterKompetensi.EnumDataType.Text Then
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = _
                                Me.ObjKompetensi.SetFormatDataCustom(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value, DataType, DefPrec, Me.ObjCekdata)
                            ElseIf DataType = ClsMasterKompetensi.EnumDataType.YN Then
                                If Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "YES" Then
                                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = "Yes"
                                ElseIf Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "NO" Then
                                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = "No"
                                ElseIf Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "Y" Then
                                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = "Yes"
                                ElseIf Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper = "N" Then
                                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = "No"
                                Else
                                    MsgBox("Must input Yes / No .", MsgBoxStyle.Information, "Information")
                                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            End If
                        End If
                    End If
                    ''sertifikasi
                    If e.ColumnIndex = Me.ColCompetence_NextCertificationDate.Index Then
                        If Me.ObjPost.GetPeriodic(Me.TxtPrimaryPosition.Text, Me.DGVKompetensi.Item(Me.ColCompetence_Primary.Index, e.RowIndex).Value) = 1 Then
                            'jika di master job position periodic update nya di ceklis maka harus input tanggal di next certification

                            'set region ke indonesia
                            My.Application.ChangeCulture("id-ID")
                            Thread.CurrentThread.CurrentCulture = New CultureInfo("id-ID", True)
                            If IsDate(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Must input date format.", MsgBoxStyle.Information, "Information")
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Else
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                            End If
                            'kembalikan ke region defaultnya
                            My.Application.ChangeCulture("en-US")
                            Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", True)
                        ElseIf Me.ObjPost.GetPeriodic(Me.TxtPrimaryPosition.Text, Me.DGVKompetensi.Item(Me.ColCompetence_Primary.Index, e.RowIndex).Value) = -1 Then
                            'set region ke indonesia
                            My.Application.ChangeCulture("id-ID")
                            Thread.CurrentThread.CurrentCulture = New CultureInfo("id-ID", True)
                            If IsDate(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                If Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = "-"
                                ElseIf Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value <> "-" Then
                                    MsgBox("Must input date format.", MsgBoxStyle.Information, "Information")
                                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            Else
                                Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                            End If
                            'kembalikan ke region defaultnya
                            My.Application.ChangeCulture("en-US")
                            Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", True)
                        End If
                    End If
                Else

                    Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        Me.mDgv = Me.DGVKompetensi
                        Me.mEraseRowIndex = e.RowIndex
                        Me.BgwDeleteRowDgv.RunWorkerAsync()
                    Else
                        Me.DGVKompetensi.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                    ObjCekdata.SetNumber(Me.DGVKompetensi)
                    Exit Sub
                End If
            End With


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVKompetensi_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DGVKompetensi.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            Me.DGVKompetensi.TabStop = True
            rowDGV = Me.DGVKompetensi.Rows(Me.DGVKompetensi.RowCount - 1)
            With rowDGV
                .Cells(Me.ColCompetence_Code.Index).Value = ""
            End With
            Dim i As Integer = 1
            For Each rowDGV In DGVKompetensi.Rows
                rowDGV.Cells(Me.ColCompetenceNo.Index).Value = i
                i = i + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVKompetensi_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DGVKompetensi.RowsRemoved
        Dim rowDGV As DataGridViewRow
        Try
            Dim i As Integer = 1
            For Each rowDGV In DGVKompetensi.Rows
                rowDGV.Cells(Me.ColCompetenceNo.Index).Value = i
                i = i + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvFamily_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvFamily.RowsRemoved
        Dim rowDGV As DataGridViewRow
        Try
            Dim i As Integer = 1
            For Each rowDGV In dgvFamily.Rows
                rowDGV.Cells(Me.ColNo.Index).Value = i
                i = i + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVKompetensi_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVKompetensi.Sorted
        Dim rowDGV As DataGridViewRow
        Try
            Dim i As Integer = 1
            For Each rowDGV In DGVKompetensi.Rows
                rowDGV.Cells(Me.ColCompetenceNo.Index).Value = i
                i = i + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvFamily_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvFamily.Sorted
        Dim rowDGV As DataGridViewRow
        Try
            Dim i As Integer = 1
            For Each rowDGV In dgvFamily.Rows
                rowDGV.Cells(Me.ColNo.Index).Value = i
                i = i + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDeleteRowDgv_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwDeleteRowDgv.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDeleteRowDgv_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwDeleteRowDgv.RunWorkerCompleted
        Try
            Me.mDgv.Rows.RemoveAt(Me.mEraseRowIndex)
            If mDgv Is Me.DGVKontrak Then
                Me.ObjCekdata.SetNumber(Me.DGVKontrak)
            ElseIf mDgv Is Me.DGVAttanchmentKontrak Then
                Me.ObjCekdata.SetNumber(Me.DGVAttanchmentKontrak)
            ElseIf mDgv Is Me.DGVAttanchment Then
                Me.ObjCekdata.SetNumber(Me.DGVAttanchment)
            End If
        
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkNoNPWP_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkNoNPWP.CheckedChanged
        Try
            If Me.ChkNoNPWP.Checked = False Then
                Me.TxtPersonal_NPWP.Text = "00.000.000.0-000.000"
                Me.TxtLegal_NPWP_Address.Text = ""
                Me.TxtPersonal_NPWP.Enabled = False
                Me.TxtLegal_NPWP_Address.Enabled = False
            Else
                Me.TxtPersonal_NPWP.Enabled = True
                Me.TxtLegal_NPWP_Address.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFingerprintID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFingerprintID.GotFocus
        Try
            If Me.TxtFingerprintID.Enabled = True Then
                If Me.TxtFingerprintID.ReadOnly = True Then
                    Me.TxtFingerprintID.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtFingerprintID.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFingerprintID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFingerprintID.LostFocus
        Try
            If Me.TxtFingerprintID.Enabled = True Then
                If Me.TxtFingerprintID.ReadOnly = True Then
                    Me.TxtFingerprintID.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtFingerprintID.BackColor = Me.TxtPrimaryEmployee.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFingerprintID_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtFingerprintID.Leave
        Try
            If String.IsNullOrEmpty(Me.TxtFingerprintID.Text) = False Then
                If Me.ObjKaryawan.IsFingerprintIDAlreadyExist(Me.TxtFingerprintID.Text, Me.TxtPrimaryEmployee.Text) = True Then
                    MsgBox("ID " & Me.TxtFingerprintID.Text & " already use.", MsgBoxStyle.Information, "Information")
                    Me.TxtFingerprintID.Focus()
                    Me.TxtFingerprintID.[Select](0, Me.TxtFingerprintID.Text.Length)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkOverrideClockIn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkOverrideClockIn.CheckedChanged
        Try
            If Me.ChkOverrideClockIn.Checked = True Then
                Me.ChkDoesnNotClockIn.Enabled = True
                Me.ChkDoesnNotClockIn.Checked = True
            ElseIf Me.ChkOverrideClockIn.Checked = False Then
                Me.ChkDoesnNotClockIn.Enabled = False
                Me.ChkDoesnNotClockIn.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkOverrideClockOut_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkOverrideClockOut.CheckedChanged
        Try
            If Me.ChkOverrideClockOut.Checked = True Then
                Me.ChkDoesnNotClockOut.Enabled = True
                Me.ChkDoesnNotClockOut.Checked = True
            ElseIf Me.ChkOverrideClockOut.Checked = False Then
                Me.ChkDoesnNotClockOut.Enabled = False
                Me.ChkDoesnNotClockOut.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkOverrideEarlyOT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkOverrideEarlyOT.CheckedChanged
        Try
            'ahdar
            If Me.TxtNIK.Enabled = True Then
                If Me.ChkOverrideEarlyOT.Checked = True Then
                    Me.CboIsCalculateEarlyOT.Enabled = True
                    If Me.CboIsCalculateEarlyOT.SelectedIndex = 0 Then
                        Me.txtCalculateAfterEarlyOT.Enabled = True
                        Me.cboRoundingTypeEarlyOT.Enabled = True
                        If Me.cboRoundingTypeEarlyOut.SelectedIndex <> 0 Then
                            Me.txtRoundingEarlyOT.Enabled = True
                        End If
                        Me.chkOverrideCalculateEarlyOTTolerance.Enabled = True
                        Me.txtMaxValueEarlyOT.Enabled = True
                    Else
                        Me.chkOverrideCalculateEarlyOTTolerance.Enabled = False
                    End If
                ElseIf Me.ChkOverrideEarlyOT.Checked = False Then
                    Me.CboIsCalculateEarlyOT.Enabled = False
                    Me.txtCalculateAfterEarlyOT.Enabled = False
                    Me.cboRoundingTypeEarlyOT.Enabled = False
                    Me.txtRoundingEarlyOT.Enabled = False
                    Me.txtMaxValueEarlyOT.Enabled = False
                    Me.txtCalculateAfterEarlyOT.Text = "00.00"
                    Me.cboRoundingTypeEarlyOT.SelectedIndex = 0
                    Me.txtRoundingEarlyOT.Text = "00.00"
                    Me.txtMaxValueEarlyOT.Text = "00.00"
                    Me.CboIsCalculateEarlyOT.SelectedIndex = 1
                End If

                Me.SetUseOvertimeOrder()
            End If
            'If Me.TxtNIK.Enabled = True Then
            '    If Me.ChkOverrideEarlyOT.Checked = True Then
            '        Me.CboIsCalculateEarlyOT.Enabled = True
            '        If Me.CboIsCalculateEarlyOT.SelectedIndex = 0 Then
            '            Me.txtCalculateAfterEarlyOT.Enabled = True
            '            Me.cboRoundingTypeEarlyOT.Enabled = True
            '            If Me.cboRoundingTypeEarlyOut.SelectedIndex <> 0 Then
            '                Me.txtRoundingEarlyOT.Enabled = True
            '            End If
            '            Me.txtMaxValueEarlyOT.Enabled = True
            '        End If
            '    ElseIf Me.ChkOverrideEarlyOT.Checked = False Then
            '        Me.CboIsCalculateEarlyOT.Enabled = False
            '        Me.txtCalculateAfterEarlyOT.Enabled = False
            '        Me.cboRoundingTypeEarlyOT.Enabled = False
            '        Me.txtRoundingEarlyOT.Enabled = False
            '        Me.txtMaxValueEarlyOT.Enabled = False
            '        Me.txtCalculateAfterEarlyOT.Text = "00.00"
            '        Me.cboRoundingTypeEarlyOT.SelectedIndex = 0
            '        Me.txtRoundingEarlyOT.Text = "00.00"
            '        Me.txtMaxValueEarlyOT.Text = "00.00"
            '        Me.CboIsCalculateEarlyOT.SelectedIndex = 1
            '    End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkOverrideLate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkOverrideLate.CheckedChanged
        Try
            'ahdar
            If Me.TxtNIK.Enabled = True Then
                If Me.ChkOverrideLate.Checked = True Then
                    Me.CboIsCalculateLate.Enabled = True
                    If Me.CboIsCalculateLate.SelectedIndex = 0 Then
                        Me.txtCalculateAfterLate.Enabled = True
                        Me.cboRoundingTypeLate.Enabled = True
                        If Me.cboRoundingTypeLate.SelectedIndex <> 0 Then
                            Me.txtRoundingLate.Enabled = True
                        End If
                        Me.txtMaxValueLate.Enabled = True
                        Me.chkOverrideCalculateLateTolerance.Enabled = True
                    Else
                        Me.chkOverrideCalculateLateTolerance.Enabled = False
                    End If
                ElseIf Me.ChkOverrideLate.Checked = False Then
                    Me.CboIsCalculateLate.Enabled = False
                    Me.txtCalculateAfterLate.Enabled = False
                    Me.cboRoundingTypeLate.Enabled = False
                    Me.txtRoundingLate.Enabled = False
                    Me.txtMaxValueLate.Enabled = False
                    Me.txtCalculateAfterLate.Text = "00.00"
                    Me.cboRoundingTypeLate.SelectedIndex = 0
                    Me.txtRoundingLate.Text = "00.00"
                    Me.txtMaxValueLate.Text = "00.00"
                    Me.CboIsCalculateLate.SelectedIndex = 1
                End If
            End If
            'If Me.TxtNIK.Enabled = True Then
            '    If Me.ChkOverrideLate.Checked = True Then
            '        Me.CboIsCalculateLate.Enabled = True
            '        If Me.CboIsCalculateLate.SelectedIndex = 0 Then
            '            Me.txtCalculateAfterLate.Enabled = True
            '            Me.cboRoundingTypeLate.Enabled = True
            '            If Me.cboRoundingTypeLate.SelectedIndex <> 0 Then
            '                Me.txtRoundingLate.Enabled = True
            '            End If
            '            Me.txtMaxValueLate.Enabled = True
            '        End If
            '    ElseIf Me.ChkOverrideLate.Checked = False Then
            '        Me.CboIsCalculateLate.Enabled = False
            '        Me.txtCalculateAfterLate.Enabled = False
            '        Me.cboRoundingTypeLate.Enabled = False
            '        Me.txtRoundingLate.Enabled = False
            '        Me.txtMaxValueLate.Enabled = False
            '        Me.txtCalculateAfterLate.Text = "00.00"
            '        Me.cboRoundingTypeLate.SelectedIndex = 0
            '        Me.txtRoundingLate.Text = "00.00"
            '        Me.txtMaxValueLate.Text = "00.00"
            '        Me.CboIsCalculateLate.SelectedIndex = 1
            '    End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkOverrideEarlyOut_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkOverrideEarlyOut.CheckedChanged
        Try
            'ahdar
            If Me.TxtNIK.Enabled = True Then
                If Me.ChkOverrideEarlyOut.Checked = True Then
                    Me.CboIsCalculateEarlyOut.Enabled = True
                    If Me.CboIsCalculateEarlyOut.SelectedIndex = 0 Then
                        Me.txtCalculateAfterEarlyOut.Enabled = True
                        Me.cboRoundingTypeEarlyOut.Enabled = True
                        If Me.cboRoundingTypeEarlyOut.SelectedIndex <> 0 Then
                            Me.txtRoundingEarlyOut.Enabled = True
                        End If
                        Me.txtMaxValueEarlyOut.Enabled = True
                        Me.chkOverrideCalculateEarlyOutTolerance.Enabled = True
                    Else
                        Me.chkOverrideCalculateEarlyOutTolerance.Enabled = False
                    End If
                ElseIf Me.ChkOverrideEarlyOut.Checked = False Then
                    Me.CboIsCalculateEarlyOut.Enabled = False
                    Me.txtCalculateAfterEarlyOut.Enabled = False
                    Me.cboRoundingTypeEarlyOut.Enabled = False
                    Me.txtRoundingEarlyOut.Enabled = False
                    Me.txtMaxValueEarlyOut.Enabled = False
                    Me.txtCalculateAfterEarlyOut.Text = "00.00"
                    Me.cboRoundingTypeEarlyOut.SelectedIndex = 0
                    Me.txtRoundingEarlyOut.Text = "00.00"
                    Me.txtMaxValueEarlyOut.Text = "00.00"
                    Me.CboIsCalculateEarlyOut.SelectedIndex = 1
                End If
            End If
            'If Me.TxtNIK.Enabled = True Then
            '    If Me.ChkOverrideEarlyOut.Checked = True Then
            '        Me.CboIsCalculateEarlyOut.Enabled = True
            '        If Me.CboIsCalculateEarlyOut.SelectedIndex = 0 Then
            '            Me.txtCalculateAfterEarlyOut.Enabled = True
            '            Me.cboRoundingTypeEarlyOut.Enabled = True
            '            If Me.cboRoundingTypeEarlyOut.SelectedIndex <> 0 Then
            '                Me.txtRoundingEarlyOut.Enabled = True
            '            End If
            '            Me.txtMaxValueEarlyOut.Enabled = True
            '        End If
            '    ElseIf Me.ChkOverrideEarlyOut.Checked = False Then
            '        Me.CboIsCalculateEarlyOut.Enabled = False
            '        Me.txtCalculateAfterEarlyOut.Enabled = False
            '        Me.cboRoundingTypeEarlyOut.Enabled = False
            '        Me.txtRoundingEarlyOut.Enabled = False
            '        Me.txtMaxValueEarlyOut.Enabled = False
            '        Me.txtCalculateAfterEarlyOut.Text = "00.00"
            '        Me.cboRoundingTypeEarlyOut.SelectedIndex = 0
            '        Me.txtRoundingEarlyOut.Text = "00.00"
            '        Me.txtMaxValueEarlyOut.Text = "00.00"
            '        Me.CboIsCalculateEarlyOut.SelectedIndex = 1
            '    End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkOverrideOT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkOverrideOT.CheckedChanged
        Try
            'ahdar
            If Me.TxtNIK.Enabled = True Then
                If Me.ChkOverrideOT.Checked = True Then
                    Me.CboIsCalculateOT.Enabled = True
                    If Me.CboIsCalculateOT.SelectedIndex = 0 Then
                        Me.txtCalculateAfterOT.Enabled = True
                        Me.cboRoundingTypeOT.Enabled = True
                        If Me.cboRoundingTypeOT.SelectedIndex <> 0 Then
                            Me.txtRoundingOT.Enabled = True
                        End If
                        Me.txtMaxValueOT.Enabled = True
                        Me.chkOverrideCalculateOTTolerance.Enabled = True
                    Else
                        Me.chkOverrideCalculateOTTolerance.Enabled = False
                    End If
                ElseIf Me.ChkOverrideOT.Checked = False Then
                    Me.CboIsCalculateOT.Enabled = False
                    Me.txtCalculateAfterOT.Enabled = False
                    Me.cboRoundingTypeOT.Enabled = False
                    Me.txtRoundingOT.Enabled = False
                    Me.txtMaxValueOT.Enabled = False
                    Me.txtCalculateAfterOT.Text = "00.00"
                    Me.cboRoundingTypeOT.SelectedIndex = 0
                    Me.txtRoundingOT.Text = "00.00"
                    Me.txtMaxValueOT.Text = "00.00"
                    Me.CboIsCalculateOT.SelectedIndex = 1
                End If

                Me.SetUseOvertimeOrder()
            End If
            'If Me.TxtNIK.Enabled = True Then
            '    If Me.ChkOverrideOT.Checked = True Then
            '        Me.CboIsCalculateOT.Enabled = True
            '        If Me.CboIsCalculateOT.SelectedIndex = 0 Then
            '            Me.txtCalculateAfterOT.Enabled = True
            '            Me.cboRoundingTypeOT.Enabled = True
            '            If Me.cboRoundingTypeOT.SelectedIndex <> 0 Then
            '                Me.txtRoundingOT.Enabled = True
            '            End If
            '            Me.txtMaxValueOT.Enabled = True
            '        End If
            '    ElseIf Me.ChkOverrideOT.Checked = False Then
            '        Me.CboIsCalculateOT.Enabled = False
            '        Me.txtCalculateAfterOT.Enabled = False
            '        Me.cboRoundingTypeOT.Enabled = False
            '        Me.txtRoundingOT.Enabled = False
            '        Me.txtMaxValueOT.Enabled = False
            '        Me.txtCalculateAfterOT.Text = "00.00"
            '        Me.cboRoundingTypeOT.SelectedIndex = 0
            '        Me.txtRoundingOT.Text = "00.00"
            '        Me.txtMaxValueOT.Text = "00.00"
            '        Me.CboIsCalculateOT.SelectedIndex = 1
            '    End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCalculateAfterEarlyOT_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCalculateAfterEarlyOT.Enter, txtCalculateAfterLate.Enter, txtCalculateAfterEarlyOut.Enter, txtCalculateAfterOT.Enter, txtRoundingEarlyOT.Enter, txtRoundingLate.Enter, txtRoundingEarlyOut.Enter, txtRoundingOT.Enter, txtMaxValueEarlyOT.Enter, txtMaxValueLate.Enter, txtMaxValueEarlyOut.Enter, txtMaxValueOT.Enter
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        Me.buffer = Txt.Text
                        Me.txtSelect = Txt
                        Me.BackgroundWorker1.RunWorkerAsync()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCalculateAfterEarlyOT_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCalculateAfterEarlyOT.GotFocus, txtCalculateAfterLate.GotFocus, txtCalculateAfterEarlyOut.GotFocus, txtCalculateAfterOT.GotFocus, txtRoundingEarlyOT.GotFocus, txtRoundingLate.GotFocus, txtRoundingEarlyOut.GotFocus, txtRoundingOT.GotFocus, txtMaxValueEarlyOT.GotFocus, txtMaxValueLate.GotFocus, txtMaxValueEarlyOut.GotFocus, txtMaxValueOT.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.BackColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCalculateAfterEarlyOT_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCalculateAfterEarlyOT.Leave, txtCalculateAfterLate.Leave, txtCalculateAfterEarlyOut.Leave, txtCalculateAfterOT.Leave, txtRoundingEarlyOT.Leave, txtRoundingLate.Leave, txtRoundingEarlyOut.Leave, txtRoundingOT.Leave, txtMaxValueEarlyOT.Leave, txtMaxValueLate.Leave, txtMaxValueEarlyOut.Leave, txtMaxValueOT.Leave
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If String.IsNullOrEmpty(Txt.Text) = False Then
                            If Me.ObjCekdata.Is_Time_String(Txt.Text) = True Then
                                If Me.ObjCekdata.Convert_ToMinutes(Txt.Text) > 1439 Then
                                    MsgBox("Schedule in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                                    Txt.Text = Me.buffer
                                End If
                            Else
                                MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
                                Txt.Text = Me.buffer
                            End If
                        Else
                            Txt.Text = "00.00"
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCalculateAfterEarlyOT_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCalculateAfterEarlyOT.LostFocus, txtCalculateAfterLate.LostFocus, txtCalculateAfterEarlyOut.LostFocus, txtCalculateAfterOT.LostFocus, txtRoundingEarlyOT.LostFocus, txtRoundingLate.LostFocus, txtRoundingEarlyOut.LostFocus, txtRoundingOT.LostFocus, txtMaxValueEarlyOT.LostFocus, txtMaxValueLate.LostFocus, txtMaxValueEarlyOut.LostFocus, txtMaxValueOT.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimaryDepartemen.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.txtSelect.SelectAll()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboIsCalculateEarlyOT_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboIsCalculateEarlyOT.SelectedIndexChanged
        Try
            If Me.TxtNIK.Enabled = True Then
                If Me.CboIsCalculateEarlyOT.SelectedIndex = 0 Then
                    Me.txtCalculateAfterEarlyOT.Enabled = True
                    Me.cboRoundingTypeEarlyOT.Enabled = True
                    If Me.cboRoundingTypeEarlyOT.SelectedIndex <> 0 Then
                        Me.txtRoundingEarlyOT.Enabled = True
                    End If
                    Me.txtMaxValueEarlyOT.Enabled = True
                    Me.chkOverrideCalculateEarlyOTTolerance.Enabled = True
                    Me.chkOverrideCalculateEarlyOTTolerance.Checked = True
                    'Me.ObjKaryawan.CalculateEarlyOTOnWeekend = 1
                    'Me.ObjKaryawan.CalculateEarlyOTOnHoliday = 1
                Else
                    Me.txtCalculateAfterEarlyOT.Enabled = False
                    Me.cboRoundingTypeEarlyOT.Enabled = False
                    Me.txtRoundingEarlyOT.Enabled = False
                    Me.txtMaxValueEarlyOT.Enabled = False
                    Me.txtCalculateAfterEarlyOT.Text = "00.00"
                    Me.cboRoundingTypeEarlyOT.SelectedIndex = 0
                    Me.txtRoundingEarlyOT.Text = "00.00"
                    Me.txtMaxValueEarlyOT.Text = "00.00"
                    Me.chkOverrideCalculateEarlyOTTolerance.Enabled = False
                    Me.chkOverrideCalculateEarlyOTTolerance.Checked = False
                    'Me.ObjKaryawan.CalculateEarlyOTOnWeekend = 0
                    'Me.ObjKaryawan.CalculateEarlyOTOnHoliday = 0
                End If

                Me.SetUseOvertimeOrder()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboIsCalculateLate_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboIsCalculateLate.SelectedIndexChanged
        Try
            If Me.TxtNIK.Enabled = True Then
                If Me.CboIsCalculateLate.SelectedIndex = 0 Then
                    Me.txtCalculateAfterLate.Enabled = True
                    Me.cboRoundingTypeLate.Enabled = True
                    If Me.cboRoundingTypeLate.SelectedIndex <> 0 Then
                        Me.txtRoundingLate.Enabled = True
                    End If
                    Me.txtMaxValueLate.Enabled = True
                    Me.chkOverrideCalculateLateTolerance.Enabled = True
                    Me.chkOverrideCalculateLateTolerance.Checked = True
                    Me.ObjKaryawan.CalculateLateOnWeekend = 1
                    Me.ObjKaryawan.CalculateLateOnHoliday = 1
                Else
                    Me.txtCalculateAfterLate.Enabled = False
                    Me.cboRoundingTypeLate.Enabled = False
                    Me.txtRoundingLate.Enabled = False
                    Me.txtMaxValueLate.Enabled = False
                    Me.txtCalculateAfterLate.Text = "00.00"
                    Me.cboRoundingTypeLate.SelectedIndex = 0
                    Me.txtRoundingLate.Text = "00.00"
                    Me.txtMaxValueLate.Text = "00.00"
                    Me.chkOverrideCalculateLateTolerance.Enabled = False
                    Me.chkOverrideCalculateLateTolerance.Checked = False
                    Me.ObjKaryawan.CalculateLateOnWeekend = 0
                    Me.ObjKaryawan.CalculateLateOnHoliday = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboIsCalculateEarlyOut_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboIsCalculateEarlyOut.SelectedIndexChanged
        Try
            If Me.TxtNIK.Enabled = True Then
                If Me.CboIsCalculateEarlyOut.SelectedIndex = 0 Then
                    Me.txtCalculateAfterEarlyOut.Enabled = True
                    Me.cboRoundingTypeEarlyOut.Enabled = True
                    If Me.cboRoundingTypeEarlyOut.SelectedIndex <> 0 Then
                        Me.txtRoundingEarlyOut.Enabled = True
                    End If
                    Me.txtMaxValueEarlyOut.Enabled = True
                    Me.chkOverrideCalculateEarlyOutTolerance.Enabled = True
                    Me.chkOverrideCalculateEarlyOutTolerance.Checked = True
                    Me.ObjKaryawan.CalculateEarlyOutOnWeekend = 1
                    Me.ObjKaryawan.CalculateEarlyOutOnHoliday = 1
                Else
                    Me.txtCalculateAfterEarlyOut.Enabled = False
                    Me.cboRoundingTypeEarlyOut.Enabled = False
                    Me.txtRoundingEarlyOut.Enabled = False
                    Me.txtMaxValueEarlyOut.Enabled = False
                    Me.txtCalculateAfterEarlyOut.Text = "00.00"
                    Me.cboRoundingTypeEarlyOut.SelectedIndex = 0
                    Me.txtRoundingEarlyOut.Text = "00.00"
                    Me.txtMaxValueEarlyOut.Text = "00.00"
                    Me.chkOverrideCalculateEarlyOutTolerance.Enabled = False
                    Me.chkOverrideCalculateEarlyOutTolerance.Checked = False
                    Me.ObjKaryawan.CalculateEarlyOutOnWeekend = 0
                    Me.ObjKaryawan.CalculateEarlyOutOnHoliday = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboIsCalculateOT_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboIsCalculateOT.SelectedIndexChanged
        Try
            If Me.TxtNIK.Enabled = True Then
                If Me.CboIsCalculateOT.SelectedIndex = 0 Then
                    Me.txtCalculateAfterOT.Enabled = True
                    Me.cboRoundingTypeOT.Enabled = True
                    If Me.cboRoundingTypeOT.SelectedIndex <> 0 Then
                        Me.txtRoundingOT.Enabled = True
                    End If
                    Me.txtMaxValueOT.Enabled = True
                    Me.chkOverrideCalculateOTTolerance.Enabled = True
                    Me.chkOverrideCalculateOTTolerance.Checked = True
                    'Me.ObjKaryawan.CalculateOTOnWeekend = 1
                    'Me.ObjKaryawan.CalculateOTOnHoliday = 1
                Else
                    Me.txtCalculateAfterOT.Enabled = False
                    Me.cboRoundingTypeOT.Enabled = False
                    Me.txtRoundingOT.Enabled = False
                    Me.txtMaxValueOT.Enabled = False
                    Me.txtCalculateAfterOT.Text = "00.00"
                    Me.cboRoundingTypeOT.SelectedIndex = 0
                    Me.txtRoundingOT.Text = "00.00"
                    Me.txtMaxValueOT.Text = "00.00"
                    Me.chkOverrideCalculateOTTolerance.Enabled = False
                    Me.chkOverrideCalculateOTTolerance.Checked = False
                    'Me.ObjKaryawan.CalculateOTOnWeekend = 0
                    'Me.ObjKaryawan.CalculateOTOnHoliday = 0
                End If

                Me.SetUseOvertimeOrder()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboRoundingTypeEarlyOT_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRoundingTypeEarlyOT.SelectedIndexChanged
        Try
            If Me.TxtNIK.Enabled = True Then
                If Me.cboRoundingTypeEarlyOT.SelectedIndex <> 0 Then
                    Me.txtRoundingEarlyOT.Enabled = True
                Else
                    Me.txtRoundingEarlyOT.Enabled = False
                    Me.txtRoundingEarlyOT.Text = "00.00"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboRoundingTypeLate_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRoundingTypeLate.SelectedIndexChanged
        Try
            If Me.TxtNIK.Enabled = True Then
                If Me.cboRoundingTypeLate.SelectedIndex <> 0 Then
                    Me.txtRoundingLate.Enabled = True
                Else
                    Me.txtRoundingLate.Enabled = False
                    Me.txtRoundingLate.Text = "00.00"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboRoundingTypeEarlyOut_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRoundingTypeEarlyOut.SelectedIndexChanged
        Try
            If Me.TxtNIK.Enabled = True Then
                If Me.cboRoundingTypeEarlyOut.SelectedIndex <> 0 Then
                    Me.txtRoundingEarlyOut.Enabled = True
                Else
                    Me.txtRoundingEarlyOut.Enabled = False
                    Me.txtRoundingEarlyOut.Text = "00.00"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboRoundingTypeOT_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRoundingTypeOT.SelectedIndexChanged
        Try
            If Me.TxtNIK.Enabled = True Then
                If Me.cboRoundingTypeOT.SelectedIndex <> 0 Then
                    Me.txtRoundingOT.Enabled = True
                Else
                    Me.txtRoundingOT.Enabled = False
                    Me.txtRoundingOT.Text = "00.00"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStatus.SelectedIndexChanged
        Try
            If Me.cboStatus.SelectedIndex = ClsMasterKaryawan.EnumAktif.Active Then
                Me.lblResignDate.Visible = False
                Me.DtpResignDate.Visible = False
                Me.chkloses_tax_obligations.Visible = False
                Me.chkloses_tax_obligations.TabStop = False
                Me.chkloses_tax_obligations.Checked = False
            Else
                Me.lblResignDate.Visible = True
                Me.DtpResignDate.Visible = True
                Me.chkloses_tax_obligations.Visible = True
                Me.chkloses_tax_obligations.TabStop = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim frmMasterWorkShift_AdvanceSettings As New FrmMasterWorkShift_AdvanceSettings
        Try
            If Me.TxtPrimaryEmployee.Text = "-1" Or (Me.CboIsCalculateEarlyOT.SelectedIndex = 1 And Me.CboIsCalculateLate.SelectedIndex = 1 And _
                    Me.CboIsCalculateEarlyOut.SelectedIndex = 1 And Me.CboIsCalculateOT.SelectedIndex = 1) Then
                Exit Sub
            End If
            frmMasterWorkShift_AdvanceSettings.Text = "Employee Attendance Advance Settings"
            frmMasterWorkShift_AdvanceSettings.CalledFromWS = False
            frmMasterWorkShift_AdvanceSettings.BackColorAll = Me.BackColorAll
            frmMasterWorkShift_AdvanceSettings.UserCode = Me.UserCode
            frmMasterWorkShift_AdvanceSettings.ObjEmployee = Me.ObjKaryawan
            If Me.CboIsCalculateEarlyOT.SelectedIndex = 1 Then
                frmMasterWorkShift_AdvanceSettings.disabledEarlyOT = True
            End If
            If Me.CboIsCalculateLate.SelectedIndex = 1 Then
                frmMasterWorkShift_AdvanceSettings.disabledLate = True
            End If
            If Me.CboIsCalculateEarlyOut.SelectedIndex = 1 Then
                frmMasterWorkShift_AdvanceSettings.disabledEarlyOut = True
            End If
            If Me.CboIsCalculateOT.SelectedIndex = 1 Then
                frmMasterWorkShift_AdvanceSettings.disabledOT = True
            End If
            If Me.btnOK.Enabled = True Then
                frmMasterWorkShift_AdvanceSettings.ReadOnly_ = False
            End If
            If frmMasterWorkShift_AdvanceSettings.ShowDialog = Windows.Forms.DialogResult.OK Then
                With Me.ObjKaryawan
                    If frmMasterWorkShift_AdvanceSettings.chkCalculateEarlyOTOnWeekend.Checked = True Then
                        .CalculateEarlyOTOnWeekend = 1
                    Else
                        .CalculateEarlyOTOnWeekend = 0
                    End If
                    If frmMasterWorkShift_AdvanceSettings.chkCalculateLateOnWeekend.Checked = True Then
                        .CalculateLateOnWeekend = 1
                    Else
                        .CalculateLateOnWeekend = 0
                    End If
                    If frmMasterWorkShift_AdvanceSettings.chkCalculateEarlyOutOnWeekend.Checked = True Then
                        .CalculateEarlyOutOnWeekend = 1
                    Else
                        .CalculateEarlyOutOnWeekend = 0
                    End If
                    If frmMasterWorkShift_AdvanceSettings.chkCalculateOTOnWeekend.Checked = True Then
                        .CalculateOTOnWeekend = 1
                    Else
                        .CalculateOTOnWeekend = 0
                    End If
                    If frmMasterWorkShift_AdvanceSettings.chkCalculateEarlyOTOnHoliday.Checked = True Then
                        .CalculateEarlyOTOnHoliday = 1
                    Else
                        .CalculateEarlyOTOnHoliday = 0
                    End If
                    If frmMasterWorkShift_AdvanceSettings.chkCalculateLateOnHoliday.Checked = True Then
                        .CalculateLateOnHoliday = 1
                    Else
                        .CalculateLateOnHoliday = 0
                    End If
                    If frmMasterWorkShift_AdvanceSettings.chkCalculateEarlyOutOnHoliday.Checked = True Then
                        .CalculateEarlyOutOnHoliday = 1
                    Else
                        .CalculateEarlyOutOnHoliday = 0
                    End If
                    If frmMasterWorkShift_AdvanceSettings.chkCalculateOTOnHoliday.Checked = True Then
                        .CalculateOTOnHoliday = 1
                    Else
                        .CalculateOTOnHoliday = 0
                    End If
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVKontrak_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVKontrak.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColContract_No.Index = True Then
                e.Cancel = True
            End If
            If e.ColumnIndex <> Me.ColContract_DocNo.Index = True Then

                If Me.DGVKontrak.Item(Me.ColContract_DocNo.Index, e.RowIndex).Value = "" Then
                    e.Cancel = True
                Else
                    Me.buffer = Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value
                End If
                If e.ColumnIndex <> Me.ColContract_Type.Index = True Then
                    If Me.DGVKontrak.Item(Me.ColContract_DocNo.Index, e.RowIndex).Value = Nothing Then
                        e.Cancel = True
                    Else
                        Me.buffer = Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    Exit Sub
                End If
                If e.ColumnIndex <> Me.ColContract_Date.Index = True Then
                    If Me.DGVKontrak.Item(Me.ColContract_Date.Index, e.RowIndex).Value = Nothing Then
                        e.Cancel = True
                    Else
                        Me.buffer = Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    Exit Sub
                End If
                If e.ColumnIndex <> Me.ColContract_Start.Index = True Then
                    If Me.DGVKontrak.Item(Me.ColContract_Start.Index, e.RowIndex).Value = Nothing Then
                        e.Cancel = True
                    Else
                        Me.buffer = Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    Exit Sub
                End If

                If e.ColumnIndex <> Me.ColContract_End.Index = True Then
                    If Me.DGVKontrak.Item(Me.ColContract_End.Index, e.RowIndex).Value = Nothing Then
                        e.Cancel = True
                    Else
                        Me.buffer = Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.buffer = Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVKontrak_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVKontrak.CellClick
        Try
            'If e.RowIndex > -1 Then
            '    reloadDataTableKontrak(Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, e.RowIndex).Value)
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVKontrak_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVKontrak.CellContentClick
       
    End Sub

    Private Sub DGVKontrak_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVKontrak.CellEndEdit
        Dim row As DataGridViewRow = Me.DGVKontrak.Rows(e.RowIndex)
        Dim Index As Integer = -1

        Dim tanggal As Date
        Try
            'Me.DGVKontrak.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'Me.DGVKontrak.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            If IsNothing(Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If Me.buffer = Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value Then
                    Exit Sub
                End If
                With row
                    If e.ColumnIndex = Me.ColContract_DocNo.Index Then
                        If Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                            Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                        End If

                        Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                        If String.IsNullOrEmpty(Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value) And Me.DGVKontrak.Rows.Count - 1 > 1 = True Then
                            Me.mDgv = Me.DGVKontrak
                            Me.mEraseRowIndex = row.Index
                            Me.BgwDeleteRowDgv.RunWorkerAsync()
                        Else
                            If ObjKaryawanType.isiCboEmployeeTypeContrack(Me.TBLKontrak) = False Then
                                MsgBox("Please input employee contract type data on master employee contract type first.", MsgBoxStyle.Information, "Information")
                                Me.mDgv = Me.DGVKontrak
                                Me.mEraseRowIndex = row.Index
                                Me.BgwDeleteRowDgv.RunWorkerAsync()
                                Exit Sub
                            Else
                                If Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                    MsgBox("Please specify a valid Document No", MsgBoxStyle.Information, "Information")
                                End If
                                Index = ObjKaryawanType.GetIndexItemKontrakTipe(DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, Me)

                                If Index < 0 Then
                                    If Me.DGVKontrak.Item(Me.ColContract_Type.Index, e.RowIndex).Value = Nothing Then
                                        cbo = Me.DGVKontrak.Rows(e.RowIndex).Cells(Me.ColContract_Type.Index)
                                        cbo.DataSource = Me.TBLKontrak.DefaultView
                                        cbo.DisplayMember = "Code"
                                        cbo.ValueMember = "Primary"

                                        RowKontrak = Me.TBLKontrak.DefaultView(0).Row
                                        Me.DGVKontrak.Item(Me.ColContract_Type.Index, e.RowIndex).Value = RowKontrak.Primary
                                        Me.DGVKontrak.Item(Me.ColContract_PrimaryType.Index, e.RowIndex).Value = RowKontrak.Primary
                                    End If

                                    tblItemTipeKontrak.DefaultView.RowFilter = "[document no]='" & Me.buffer & "' AND [identifier kontrak] = '" & Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, e.RowIndex).Value & "'"
                                    If tblItemTipeKontrak.DefaultView.Count > 0 Then
                                        rowItemTipeKontrak = tblItemTipeKontrak.DefaultView(0).Row
                                        rowItemTipeKontrak.document_no = Me.DGVKontrak.Item(Me.ColContract_DocNo.Index, e.RowIndex).Value
                                    Else
                                        Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, e.RowIndex).Value = Me.ObjKaryawanType.GetIdentifierBaru(Me)
                                        Me.ObjKaryawanType.TambahBarisDGVKontrak(Me.tblItemTipeKontrak, Me.DGVKontrak.Item(Me.ColContract_PrimaryType.Index, e.RowIndex).Value, _
                                        Me.TxtPrimaryEmployee.Text, Me.DGVKontrak.Item(Me.ColContract_DocNo.Index, e.RowIndex).Value, _
                                        Me.DGVKontrak.Item(Me.ColContract_Date.Index, e.RowIndex).Value, _
                                        Me.DGVKontrak.Item(Me.ColContract_Start.Index, e.RowIndex).Value, _
                                        Me.DGVKontrak.Item(Me.ColContract_End.Index, e.RowIndex).Value, _
                                        Me.DGVKontrak.Item(Me.ColContract_Remarks.Index, e.RowIndex).Value, _
                                        Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, e.RowIndex).Value, Me)
                                    End If
                                Else
                                    If e.RowIndex <> Index Then
                                        MsgBox("Duplication document no, please input another document.", MsgBoxStyle.Information, "Information")
                                        Me.DGVKontrak.Item(Me.ColContract_DocNo.Index, e.RowIndex).Value = Nothing
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.mDgv = Me.DGVKontrak
                                            Me.mEraseRowIndex = e.RowIndex
                                            Me.BgwDeleteRowDgv.RunWorkerAsync()
                                        Else
                                            Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCekdata.SetNumber(Me.DGVKontrak)
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If

                    If e.ColumnIndex = Me.ColContract_Type.Index Then
                        Me.DGVKontrak.Item(Me.ColContract_PrimaryType.Index, e.RowIndex).Value = Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value
                        tblItemTipeKontrak.DefaultView.RowFilter = "[primary tipe kontrak]= " & Me.buffer & " AND [identifier kontrak] = '" & Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, e.RowIndex).Value & "'"
                        If tblItemTipeKontrak.DefaultView.Count > 0 Then
                            rowItemTipeKontrak = tblItemTipeKontrak.DefaultView(0).Row
                            rowItemTipeKontrak.primary_tipe_kontrak = Me.DGVKontrak.Item(Me.ColContract_PrimaryType.Index, e.RowIndex).Value
                        End If
                    End If

                    If e.ColumnIndex = Me.ColContract_Date.Index Or e.ColumnIndex = Me.ColContract_Start.Index Or e.ColumnIndex = Me.ColContract_End.Index Then

                        If IsDate(Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Must input date format.", MsgBoxStyle.Information, "Information")
                            If Me.buffer = "" Then
                                Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Else
                                Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value = CDate(Me.buffer)
                            End If

                            Exit Sub
                        Else
                            tanggal = CDate(Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value = tanggal
                        End If

                        If Me.DGVKontrak.Item(Me.ColContract_End.Index, e.RowIndex).Value.ToString <> "" Then
                            Dim starDate As Date = Me.DGVKontrak.Item(Me.ColContract_Start.Index, e.RowIndex).Value
                            Dim endDate As Date = Me.DGVKontrak.Item(Me.ColContract_End.Index, e.RowIndex).Value
                            If starDate > endDate Then
                                MsgBox("Start date must be < End date", MsgBoxStyle.Information, "Information")
                                If e.ColumnIndex = Me.ColContract_Start.Index Then
                                    If Me.buffer = "" Then
                                        Me.DGVKontrak.Item(Me.ColContract_Start.Index, e.RowIndex).Value = Me.buffer
                                    Else
                                        Me.DGVKontrak.Item(Me.ColContract_Start.Index, e.RowIndex).Value = CDate(Me.buffer)
                                    End If
                                ElseIf e.ColumnIndex = Me.ColContract_End.Index Then
                                    If Me.buffer = "" Then
                                        Me.DGVKontrak.Item(Me.ColContract_End.Index, e.RowIndex).Value = Me.buffer
                                    Else
                                        Me.DGVKontrak.Item(Me.ColContract_End.Index, e.RowIndex).Value = CDate(Me.buffer)
                                    End If
                                End If
                                Exit Sub
                            Else
                                If e.ColumnIndex = Me.ColContract_Start.Index Or e.ColumnIndex = Me.ColContract_End.Index Then
                                    If Me.ObjKaryawanType.CekValueDate(Me) = False Then
                                        If Me.buffer = "" Then
                                            Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Else
                                            Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value = CDate(Me.buffer)
                                        End If
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If

                        If Me.buffer = "" Then
                            Me.buffer = Format(Now, "dd MMM yyyy").ToString 'Now 
                        End If
                        If e.ColumnIndex = Me.ColContract_Date.Index Then
                            tblItemTipeKontrak.DefaultView.RowFilter = "[date]=#" & Me.buffer & "# AND [identifier kontrak] = '" & Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, e.RowIndex).Value & "'"
                            If tblItemTipeKontrak.DefaultView.Count > 0 Then
                                rowItemTipeKontrak = tblItemTipeKontrak.DefaultView(0).Row
                                rowItemTipeKontrak._date = tanggal
                            End If
                        ElseIf e.ColumnIndex = Me.ColContract_Start.Index Then
                            tblItemTipeKontrak.DefaultView.RowFilter = "[start date]=#" & Me.buffer & "# AND [identifier kontrak] = '" & Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, e.RowIndex).Value & "'"
                            If tblItemTipeKontrak.DefaultView.Count > 0 Then
                                rowItemTipeKontrak = tblItemTipeKontrak.DefaultView(0).Row
                                rowItemTipeKontrak.start_date = tanggal
                            End If
                        ElseIf e.ColumnIndex = Me.ColContract_End.Index Then
                            tblItemTipeKontrak.DefaultView.RowFilter = "[end date]=#" & Me.buffer & "# AND [identifier kontrak] = '" & Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, e.RowIndex).Value & "'"
                            If tblItemTipeKontrak.DefaultView.Count > 0 Then
                                rowItemTipeKontrak = tblItemTipeKontrak.DefaultView(0).Row
                                rowItemTipeKontrak.end_date = tanggal
                            End If
                        End If
                    End If
                    If e.ColumnIndex = Me.ColContract_Remarks.Index Then
                        'max 250 karakter
                        If Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                            Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                        End If
                        Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("'", "") _
                        .Replace("`", "").Replace("%", "").Replace("*", "")

                        tblItemTipeKontrak.DefaultView.RowFilter = "[remarks]='" & Me.buffer & "' AND [identifier kontrak] = '" & Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, e.RowIndex).Value & "'"
                        If tblItemTipeKontrak.DefaultView.Count > 0 Then
                            rowItemTipeKontrak = tblItemTipeKontrak.DefaultView(0).Row
                            rowItemTipeKontrak.remarks = Me.DGVKontrak.Item(Me.ColContract_Remarks.Index, e.RowIndex).Value
                        End If
                    End If
                End With
            Else

                If e.ColumnIndex = Me.ColContract_DocNo.Index Then
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        Me.mDgv = Me.DGVKontrak
                        Me.mEraseRowIndex = row.Index
                        Me.BgwDeleteRowDgv.RunWorkerAsync()
                    Else
                        Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                Else
                    If e.ColumnIndex = Me.ColContract_Remarks.Index Then
                        Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        tblItemTipeKontrak.DefaultView.RowFilter = "[remarks]='" & Me.buffer & "' AND [identifier kontrak] = '" & Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, e.RowIndex).Value & "'"
                        If tblItemTipeKontrak.DefaultView.Count > 0 Then
                            rowItemTipeKontrak = tblItemTipeKontrak.DefaultView(0).Row
                            rowItemTipeKontrak.remarks = Me.DGVKontrak.Item(Me.ColContract_Remarks.Index, e.RowIndex).Value
                        End If
                    Else
                        Me.DGVKontrak.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If

                End If
                ObjCekdata.SetNumber(Me.DGVKontrak)
                Exit Sub
            End If
            'Me.DGVKontrak.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'Me.DGVKontrak.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            'Finally
            '    
        End Try
    End Sub

   
    Private Sub DGVKontrak_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DGVKontrak.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try

            Me.DGVKontrak.TabStop = True
            rowDGV = Me.DGVKontrak.Rows(Me.DGVKontrak.RowCount - 1)
            With rowDGV
                .Cells(Me.ColContract_DocNo.Index).Value = ""
                .Cells(Me.ColContract_Date.Index).Value = ""
                .Cells(Me.ColContract_Start.Index).Value = ""
                .Cells(Me.ColContract_End.Index).Value = ""
                .Cells(Me.ColContract_Remarks.Index).Value = ""
                .Cells(Me.ColContract_PrimaryType.Index).Value = -1
                .Cells(Me.ColContract_IdentifierContract.Index).Value = "-1-0"
            End With
            Me.ObjCekdata.SetNumber(Me.DGVKontrak)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttanchmentKontrak_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVAttanchmentKontrak.CellBeginEdit
        Try
            If Me.DGVKontrak.Rows.Count - 1 = 0 Then
                e.Cancel = True
            Else
                If Me.DGVKontrak.Item(Me.ColContract_DocNo.Index, Me.DGVKontrak.CurrentRow.Index).Value = "" Then
                    e.Cancel = True
                End If
            End If

            'jika baris baru, maka yang bisa diedit hanya desc-nya saja
            If Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Description.Index, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColKontrakAtt_Description.Index Then
                    Me.buffer = Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                'jik bukan baris baru, yang bisa diedit adalah:
                'desc, remarks, file path, 
                If e.ColumnIndex = Me.ColKontrakAtt_Description.Index Or _
                e.ColumnIndex = Me.ColKontrakAtt_LocationFile.Index Or _
                e.ColumnIndex = Me.ColKontrakAtt_Remarks.Index Then
                    Me.buffer = Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttanchmentKontrak_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVAttanchmentKontrak.CellClick
        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Description.Index, e.RowIndex).Value <> "" Then
                If e.ColumnIndex = Me.ColKontrakAtt_Preview.Index Then
                    If Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_LocationFile.Index, e.RowIndex).Value <> "" Then
                        If Me.ObjCekdata.ApakahUrlValid(Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_LocationFile.Index, e.RowIndex).Value) = True Then
                            'jika valid url maka tampilkan ke browser
                            Process.Start(Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_LocationFile.Index, e.RowIndex).Value)
                        Else
                            'jika file path, maka buka filenya
                            Dim fileExists As Boolean
                            fileExists = My.Computer.FileSystem.FileExists(Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_LocationFile.Index, e.RowIndex).Value)
                            If fileExists = False Then
                                MsgBox("You are trying to open an inaccessible file or the file does not exist anymore.", MsgBoxStyle.Information, "Information")
                            Else
                                Process.Start(Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_LocationFile.Index, e.RowIndex).Value)
                            End If
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColKontrakAtt_BtnOpenDelete.Index Then
                    If Me.btnOK.Enabled = True Then
                        Me.buffer = Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_LocationFile.Index, e.RowIndex).Value
                        If Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_BtnOpenDelete.Index, e.RowIndex).Value = "add" Then
                            'add file
                            If Me.OpenFileDialog2.ShowDialog = Windows.Forms.DialogResult.OK Then
                                If Me.OpenFileDialog2.FileName <> "" Then
                                    Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_LocationFile.Index, e.RowIndex).Value = Me.OpenFileDialog2.FileName
                                    Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Preview.Index, e.RowIndex).Value = "preview"
                                    Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_BtnOpenDelete.Index, e.RowIndex).Value = "remove"

                                End If
                            End If

                        Else
                            'delete file
                            If MsgBox("Remove file path?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_LocationFile.Index, e.RowIndex).Value = ""
                                Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Preview.Index, e.RowIndex).Value = "-"
                                Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_BtnOpenDelete.Index, e.RowIndex).Value = "add"
                            End If

                        End If
                        tblItemAttch.DefaultView.RowFilter = "[identifier kontrak]='" & Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, Me.DGVKontrak.CurrentRow.Index).Value & "' AND [desc] = '" & _
                                                            Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Description.Index, e.RowIndex).Value & "'"
                        If tblItemAttch.DefaultView.Count > 0 Then
                            rowItemAttch = tblItemAttch.DefaultView(0).Row
                            rowItemAttch.file_path = Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_LocationFile.Index, e.RowIndex).Value
                        End If

                    End If
                End If
            End If
            Me.ObjCekdata.SetNumber(Me.DGVAttanchmentKontrak)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttanchmentKontrak_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVAttanchmentKontrak.CellEndEdit
        Dim index As Integer = -1
        Try
            Dim row As DataGridViewRow = Me.DGVAttanchmentKontrak.Rows(e.RowIndex)
            Me.DGVAttanchmentKontrak.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DGVAttanchmentKontrak.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            With row
                If IsNothing(Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.buffer = Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value _
                    .ToString.Trim
                    If e.ColumnIndex = Me.ColKontrakAtt_Description.Index Then
                        If Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                            Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                        End If
                        'If Me.DGVAttanchmentKontrak.Rows.Count - 1 > 1 Then
                        '    If Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_LocationFile.Index, Me.DGVAttanchmentKontrak.Rows.Count - 3).Value = "" Then
                        '        MsgBox("Please specify a valid data in one row", MsgBoxStyle.Information, "Information")
                        '        Me.mDgv = Me.DGVAttanchmentKontrak
                        '        Me.mEraseRowIndex = row.Index
                        '        Me.BgwDeleteRowDgv.RunWorkerAsync()
                        '        Exit Sub
                        '    End If
                        'End If

                        'pastikan ada isinya
                        If Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                            MsgBox("Please specify a valid description.", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.mDgv = Me.DGVAttanchmentKontrak
                                Me.mEraseRowIndex = e.RowIndex
                                Me.BgwDeleteRowDgv.RunWorkerAsync()
                            Else
                                Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Exit Sub
                        End If
                        'upper case
                        Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'pastikan tidak ad duplikasi
                        index = GetIndex_ItemAttachment(Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Description.Index, e.RowIndex).Value, e.RowIndex)
                        If index >= 0 Then
                            MsgBox("Duplication in description data, please input another description.", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.mDgv = Me.DGVAttanchmentKontrak
                                Me.mEraseRowIndex = e.RowIndex
                                Me.BgwDeleteRowDgv.RunWorkerAsync()
                            Else
                                Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Exit Sub
                        Else
                            tblItemAttch.DefaultView.RowFilter = "[desc]='" & Me.buffer & "'"
                            If tblItemAttch.DefaultView.Count > 0 Then
                                rowItemAttch = tblItemAttch.DefaultView(0).Row
                                rowItemAttch.desc = Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Description.Index, e.RowIndex).Value
                            Else
                                Me.ObjKaryawanType.TambahBarisDGVAttchKontrak(Me.tblItemAttch, Me.TxtPrimaryEmployee.Text, _
                                Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Description.Index, e.RowIndex).Value, _
                                Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_LocationFile.Index, e.RowIndex).Value, _
                                Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Remarks.Index, e.RowIndex).Value, _
                                Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, Me.DGVKontrak.CurrentRow.Index).Value, Me)
                            End If
                        End If

                    ElseIf e.ColumnIndex = Me.ColKontrakAtt_Remarks.Index Then
                        'max 250 karakter
                        If Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                            Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                        End If
                        Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("'", "") _
                        .Replace("`", "").Replace("%", "").Replace("*", "")

                        tblItemAttch.DefaultView.RowFilter = "[remarks]='" & Me.buffer & "' AND [desc]='" & Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Description.Index, e.RowIndex).Value & "'"
                        If tblItemAttch.DefaultView.Count > 0 Then
                            rowItemAttch = tblItemAttch.DefaultView(0).Row
                            rowItemAttch.remarks = Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Remarks.Index, e.RowIndex).Value
                        End If
                    ElseIf e.ColumnIndex = Me.ColKontrakAtt_LocationFile.Index Then
                        'max 250 karakter
                        If Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                            Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                        End If
                        'jika bukan URL, maka pastikan filenya exist
                        If Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value.ToString <> "" Then
                            If Me.ObjCekdata.ApakahUrlValid(Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                If My.Computer.FileSystem.FileExists((Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value)) = False Then
                                    MsgBox("The file path you specified is not a valid url path nor a valid file path.", MsgBoxStyle.Information, "Information")
                                    Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    If Me.buffer = "" Then
                                        Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Preview.Index, e.RowIndex).Value = "-"
                                        Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_BtnOpenDelete.Index, e.RowIndex).Value = "add"
                                    End If
                                    Exit Sub
                                Else
                                    Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Preview.Index, e.RowIndex).Value = "preview"
                                    Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_BtnOpenDelete.Index, e.RowIndex).Value = "remove"
                                End If
                            Else
                                Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Preview.Index, e.RowIndex).Value = "preview"
                                Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_BtnOpenDelete.Index, e.RowIndex).Value = "remove"
                            End If
                            tblItemAttch.DefaultView.RowFilter = "[file path]='" & Me.buffer & "' AND [desc]='" & Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Description.Index, e.RowIndex).Value & "'"
                            If tblItemAttch.DefaultView.Count > 0 Then
                                rowItemAttch = tblItemAttch.DefaultView(0).Row
                                rowItemAttch.file_path = Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_LocationFile.Index, e.RowIndex).Value
                            End If
                        Else
                            Me.DGVAttanchment.Item(Me.ColKontrakAtt_Preview.Index, e.RowIndex).Value = "-"
                            Me.DGVAttanchment.Item(Me.ColKontrakAtt_BtnOpenDelete.Index, e.RowIndex).Value = "add"
                        End If

                    End If
                Else
                    If e.ColumnIndex = Me.ColKontrakAtt_Description.Index Then
                        If String.IsNullOrEmpty(Me.buffer) = True Then
                            Me.mDgv = Me.DGVAttanchmentKontrak
                            Me.mEraseRowIndex = row.Index
                            Me.BgwDeleteRowDgv.RunWorkerAsync()
                        Else
                            Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    ElseIf e.ColumnIndex = Me.ColKontrakAtt_LocationFile.Index Then
                        If Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                            Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    ElseIf e.ColumnIndex = Me.ColKontrakAtt_Remarks.Index Then

                        Me.DGVAttanchmentKontrak.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        tblItemAttch.DefaultView.RowFilter = "[remarks]='" & Me.buffer & "' AND [desc]='" & Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Description.Index, e.RowIndex).Value & "'"
                        If tblItemAttch.DefaultView.Count > 0 Then
                            rowItemAttch = tblItemAttch.DefaultView(0).Row
                            rowItemAttch.remarks = Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Remarks.Index, e.RowIndex).Value
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DGVAttanchmentKontrak.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DGVAttanchmentKontrak.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        End Try
    End Sub
    Private Function GetIndex_ItemAttachment(ByVal vDesc As String, ByVal vIndexAcuan As Integer) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.DGVAttanchmentKontrak.Rows
                If rowDGV.Index = vIndexAcuan Or rowDGV.Cells(Me.ColKontrakAtt_Description.Index).Value = "" Then
                Else
                    If rowDGV.Cells(Me.ColKontrakAtt_Description.Index).Value = vDesc Then
                        Return rowDGV.Index
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function
    Private Function GetIndex_Attachment(ByVal vDesc As String, ByVal vIndexAcuan As Integer) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.DGVAttanchment.Rows
                If rowDGV.Index = vIndexAcuan Or rowDGV.Cells(Me.ColAtt_Description.Index).Value = "" Then
                Else
                    If rowDGV.Cells(Me.ColAtt_Description.Index).Value = vDesc Then
                        Return rowDGV.Index
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function
    Private Sub DGVAttanchmentKontrak_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DGVAttanchmentKontrak.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            If Me.btnOK.Enabled = True Then
                rowDGV = Me.DGVAttanchmentKontrak.Rows(Me.DGVAttanchmentKontrak.Rows.Count - 1)
                With rowDGV

                    .Cells(Me.ColKontrakAtt_Description.Index).Value = ""
                    .Cells(Me.ColKontrakAtt_LocationFile.Index).Value = ""
                    .Cells(Me.ColKontrakAtt_No.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColKontrakAtt_BtnOpenDelete.Index).Value = "add"
                    .Cells(Me.ColKontrakAtt_Preview.Index).Value = "-"
                    .Cells(Me.ColKontrakAtt_Remarks.Index).Value = ""
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttanchment_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVAttanchment.CellBeginEdit
        Try
            'jika baris baru, maka yang bisa diedit hanya desc-nya saja
            If Me.DGVAttanchment.Item(Me.ColAtt_Description.Index, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColAtt_Description.Index Then
                    Me.buffer = Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                'jik bukan baris baru, yang bisa diedit adalah:
                'desc, remarks, file path, 
                If e.ColumnIndex = Me.ColAtt_Description.Index Or _
                e.ColumnIndex = Me.ColAtt_LocationFile.Index Or _
                e.ColumnIndex = Me.ColAtt_Remarks.Index Then
                    Me.buffer = Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttanchment_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVAttanchment.CellClick
        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If Me.DGVAttanchment.Item(Me.ColAtt_Description.Index, e.RowIndex).Value <> "" Then
                If e.ColumnIndex = Me.ColAtt_Preview.Index Then
                    If Me.DGVAttanchment.Item(Me.ColAtt_LocationFile.Index, e.RowIndex).Value <> "" Then
                        If Me.ObjCekdata.ApakahUrlValid(Me.DGVAttanchment.Item(Me.ColAtt_LocationFile.Index, e.RowIndex).Value) = True Then
                            'jika valid url maka tampilkan ke browser
                            Process.Start(Me.DGVAttanchment.Item(Me.ColAtt_LocationFile.Index, e.RowIndex).Value)
                        Else
                            'jika file path, maka buka filenya
                            Dim fileExists As Boolean
                            fileExists = My.Computer.FileSystem.FileExists(Me.DGVAttanchment.Item(Me.ColAtt_LocationFile.Index, e.RowIndex).Value)
                            If fileExists = False Then
                                MsgBox("You are trying to open an inaccessible file or the file does not exist anymore.", MsgBoxStyle.Information, "Information")
                            Else
                                Process.Start(Me.DGVAttanchment.Item(Me.ColAtt_LocationFile.Index, e.RowIndex).Value)
                            End If
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColAtt_BtnOpenDelete.Index Then
                    If Me.btnOK.Enabled = True Then
                        Me.buffer = Me.DGVAttanchment.Item(Me.ColAtt_LocationFile.Index, e.RowIndex).Value
                        If Me.DGVAttanchment.Item(Me.ColAtt_BtnOpenDelete.Index, e.RowIndex).Value = "add" Then
                            'add file
                            If Me.OpenFileDialog2.ShowDialog = Windows.Forms.DialogResult.OK Then
                                If Me.OpenFileDialog2.FileName <> "" Then
                                    Me.DGVAttanchment.Item(Me.ColAtt_LocationFile.Index, e.RowIndex).Value = Me.OpenFileDialog2.FileName
                                    Me.DGVAttanchment.Item(Me.ColAtt_Preview.Index, e.RowIndex).Value = "preview"
                                    Me.DGVAttanchment.Item(Me.ColAtt_BtnOpenDelete.Index, e.RowIndex).Value = "remove"

                                End If
                            End If

                        Else
                            'delete file
                            If MsgBox("Remove file path?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                Me.DGVAttanchment.Item(Me.ColAtt_LocationFile.Index, e.RowIndex).Value = ""
                                Me.DGVAttanchment.Item(Me.ColAtt_Preview.Index, e.RowIndex).Value = "-"
                                Me.DGVAttanchment.Item(Me.ColAtt_BtnOpenDelete.Index, e.RowIndex).Value = "add"
                            End If

                        End If
                    End If
                End If
            End If
            Me.ObjCekdata.SetNumber(Me.DGVAttanchment)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttanchment_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVAttanchment.CellEndEdit
        Dim index As Integer = -1
        Try
            Dim row As DataGridViewRow = Me.DGVAttanchment.Rows(e.RowIndex)
            Me.DGVAttanchment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DGVAttanchment.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            With row
                If IsNothing(Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.buffer = Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value _
                    .ToString.Trim
                    If e.ColumnIndex = Me.ColAtt_Description.Index Then
                        If Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                            Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                        End If
                        'pastikan ada isinya
                        If Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                            MsgBox("Please specify a valid description.", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.mDgv = Me.DGVAttanchment
                                Me.mEraseRowIndex = e.RowIndex
                                Me.BgwDeleteRowDgv.RunWorkerAsync()
                            Else
                                Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Exit Sub
                        End If
                        'upper case
                        Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'pastikan tidak ad duplikasi
                        index = GetIndex_Attachment(Me.DGVAttanchment.Item(Me.ColAtt_Description.Index, e.RowIndex).Value, e.RowIndex)
                        If index >= 0 Then
                            MsgBox("Duplication in description data, please input another description.", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.mDgv = Me.DGVAttanchment
                                Me.mEraseRowIndex = e.RowIndex
                                Me.BgwDeleteRowDgv.RunWorkerAsync()
                            Else
                                Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Exit Sub
                        End If

                    ElseIf e.ColumnIndex = Me.ColKontrakAtt_Remarks.Index Then
                        'max 250 karakter
                        If Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                            Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                        End If
                    ElseIf e.ColumnIndex = Me.ColAtt_LocationFile.Index Then
                        'max 250 karakter
                        If Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                            Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                        End If
                        'jika bukan URL, maka pastikan filenya exist
                        If Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value.ToString <> "" Then
                            If Me.ObjCekdata.ApakahUrlValid(Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                If My.Computer.FileSystem.FileExists((Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value)) = False Then
                                    MsgBox("The file path you specified is not a valid url path nor a valid file path.", MsgBoxStyle.Information, "Information")
                                    Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    If Me.buffer = "" Then
                                        Me.DGVAttanchment.Item(Me.ColAtt_Preview.Index, e.RowIndex).Value = "-"
                                        Me.DGVAttanchment.Item(Me.ColAtt_BtnOpenDelete.Index, e.RowIndex).Value = "add"
                                    End If
                                    Exit Sub
                                Else
                                    Me.DGVAttanchment.Item(Me.ColAtt_Preview.Index, e.RowIndex).Value = "preview"
                                    Me.DGVAttanchment.Item(Me.ColAtt_BtnOpenDelete.Index, e.RowIndex).Value = "remove"
                                End If
                            Else
                                Me.DGVAttanchment.Item(Me.ColAtt_Preview.Index, e.RowIndex).Value = "preview"
                                Me.DGVAttanchment.Item(Me.ColAtt_BtnOpenDelete.Index, e.RowIndex).Value = "remove"
                            End If
                        Else
                            Me.DGVAttanchment.Item(Me.ColAtt_Preview.Index, e.RowIndex).Value = "-"
                            Me.DGVAttanchment.Item(Me.ColAtt_BtnOpenDelete.Index, e.RowIndex).Value = "add"
                        End If
                    ElseIf e.ColumnIndex = Me.ColAtt_Remarks.Index Then
                        Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("'", "") _
                        .Replace("`", "").Replace("%", "").Replace("*", "")

                    End If
                Else
                    If e.ColumnIndex = Me.ColAtt_Description.Index Then
                        If String.IsNullOrEmpty(Me.buffer) = True Then
                            Me.mDgv = Me.DGVAttanchment
                            Me.mEraseRowIndex = row.Index
                            Me.BgwDeleteRowDgv.RunWorkerAsync()
                        Else
                            Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    ElseIf e.ColumnIndex = Me.ColAtt_LocationFile.Index Then
                        If Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                            Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    ElseIf e.ColumnIndex = Me.ColKontrakAtt_Remarks.Index Then
                        Me.DGVAttanchment.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DGVAttanchment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DGVAttanchment.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        End Try
    End Sub

    Private Sub DGVAttanchment_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DGVAttanchment.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            If Me.btnOK.Enabled = True Then
                rowDGV = Me.DGVAttanchment.Rows(Me.DGVAttanchment.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColAtt_Description.Index).Value = ""
                    .Cells(Me.ColAtt_LocationFile.Index).Value = ""
                    .Cells(Me.ColAtt_BtnOpenDelete.Index).Value = "add"
                    .Cells(Me.ColAtt_Preview.Index).Value = "-"
                    .Cells(Me.ColAtt_Remarks.Index).Value = ""
                End With
            End If
            Me.ObjCekdata.SetNumber(Me.DGVAttanchment)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

   
    End Sub

    Private Sub DGVKontrak_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVKontrak.SelectionChanged
        Try
            'If Me.DGVKontrak.AllowUserToAddRows = True Then
            If Me.DGVKontrak.SelectedCells.Count > 0 Then
                reloadDataTableKontrak(Me.DGVKontrak.Rows(Me.DGVKontrak.SelectedCells(0).RowIndex).Cells(Me.ColContract_IdentifierContract.Index).Value)
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVKontrak_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVKontrak.Sorted
        Dim RowDGV As DataGridViewRow
        Try

            'Exit Sub
            'Me.tblItemTipeKontrak.Clear()
            'For Each RowDGV In Me.DGVKontrak.Rows
            '    If RowDGV.Cells(Me.ColContract_DocNo.Index).Value <> "" Then
            '        Me.ObjKaryawanType.TambahBarisDGVKontrak(Me.tblItemTipeKontrak, RowDGV.Cells(Me.ColContract_PrimaryType.Index).Value, _
            '        Me.TxtPrimaryEmployee.Text, RowDGV.Cells(Me.ColContract_DocNo.Index).Value, _
            '        RowDGV.Cells(Me.ColContract_Date.Index).Value, _
            '        RowDGV.Cells(Me.ColContract_Start.Index).Value, _
            '        RowDGV.Cells(Me.ColContract_End.Index).Value, _
            '        RowDGV.Cells(Me.ColContract_Remarks.Index).Value, _
            '        RowDGV.Cells(Me.ColContract_IdentifierContract.Index).Value, Me)
            '    End If
            'Next

            'Me.ObjCekdata.SetNumber(Me.DGVKontrak)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVKontrak_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DGVKontrak.UserDeletedRow
        Try
            Me.ObjCekdata.SetNumber(Me.DGVKontrak)
            Me.DGVKontrak.Rows(0).Selected = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttanchmentKontrak_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVAttanchmentKontrak.Sorted
        Dim RowDGV As DataGridViewRow


        Try
            If Me.DGVKontrak.Rows.Count > 0 Then
                tblItemAttch.DefaultView.RowFilter = "[identifier kontrak] = '" & Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, Me.DGVKontrak.CurrentRow.Index).Value & "'"
                If tblItemAttch.DefaultView.Count > 0 Then
                    While tblItemAttch.DefaultView.Count > 0
                        rowItemAttch = tblItemAttch.DefaultView(0).Row
                        tblItemAttch.Rows.Remove(rowItemAttch)
                    End While
                End If
            End If

            For Each RowDGV In Me.DGVAttanchmentKontrak.Rows
                If RowDGV.Cells(Me.ColKontrakAtt_Description.Index).Value <> "" Then
                    Me.ObjKaryawanType.TambahBarisDGVAttchKontrak(Me.tblItemAttch, Me.TxtPrimaryEmployee.Text, _
                    RowDGV.Cells(Me.ColKontrakAtt_Description.Index).Value, _
                    RowDGV.Cells(Me.ColKontrakAtt_LocationFile.Index).Value, _
                    RowDGV.Cells(Me.ColKontrakAtt_Remarks.Index).Value, _
                    Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, Me.DGVKontrak.CurrentRow.Index).Value, Me)
                End If
            Next
            Me.ObjCekdata.SetNumber(Me.DGVAttanchmentKontrak)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttanchmentKontrak_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DGVAttanchmentKontrak.UserDeletedRow
        Try
            Me.ObjCekdata.SetNumber(Me.DGVAttanchmentKontrak)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Sub reloadDataTableKontrak(ByVal vIdentifier As String)
        Try
            Dim dr As DataGridViewRow
            tblItemAttch.DefaultView.RowFilter = "[identifier kontrak] = '" & vIdentifier & "'"
            Me.DGVAttanchmentKontrak.Rows.Clear()
            Me.DGVAttanchmentKontrak.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DGVAttanchmentKontrak.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DGVAttanchmentKontrak.AllowUserToAddRows = False
            For i As Integer = 0 To tblItemAttch.DefaultView.Count - 1
                rowItemAttch = tblItemAttch.DefaultView(i).Row
                DGVAttanchmentKontrak.Rows.Add()
                dr = DGVAttanchmentKontrak.Rows(DGVAttanchmentKontrak.Rows.Count - 1)
                dr.Cells(Me.ColKontrakAtt_No.Index).Value = Me.DGVAttanchmentKontrak.Rows.Count
                dr.Cells(Me.ColKontrakAtt_Description.Index).Value = rowItemAttch.desc
                dr.Cells(Me.ColKontrakAtt_LocationFile.Index).Value = rowItemAttch.file_path
                dr.Cells(Me.ColKontrakAtt_Remarks.Index).Value = rowItemAttch.remarks
                If rowItemAttch.file_path <> "" Then
                    dr.Cells(Me.ColKontrakAtt_Preview.Index).Value = "preview"
                    dr.Cells(Me.ColKontrakAtt_BtnOpenDelete.Index).Value = "remove"
                Else
                    dr.Cells(Me.ColKontrakAtt_Preview.Index).Value = "-"
                    dr.Cells(Me.ColKontrakAtt_BtnOpenDelete.Index).Value = "add"
                End If
            Next
            Me.DGVAttanchmentKontrak.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DGVAttanchmentKontrak.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            If Me.btnOK.Enabled = True Then
                Me.DGVAttanchmentKontrak.AllowUserToAddRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub DGVKontrak_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DGVKontrak.UserDeletingRow
        Dim identifier As String
        Try
            identifier = Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, Me.DGVKontrak.CurrentRow.Index).Value
            Me.tblItemAttch.DefaultView.RowFilter = "[identifier kontrak] = '" & identifier & "'"
            For x As Integer = 0 To Me.tblItemAttch.DefaultView.Count - 1
                If Me.tblItemAttch.DefaultView.Count > 0 Then
                    Me.rowItemAttch = Me.tblItemAttch.DefaultView(0).Row
                    Me.tblItemAttch.Rows.Remove(Me.rowItemAttch)
                End If
            Next
            identifier = Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, e.Row.Index).Value
            Me.tblItemTipeKontrak.DefaultView.RowFilter = "[identifier kontrak] = '" & identifier & "'"
            If Me.tblItemTipeKontrak.DefaultView.Count > 0 Then
                Me.rowItemTipeKontrak = Me.tblItemTipeKontrak.DefaultView(0).Row
                Me.tblItemTipeKontrak.Rows.Remove(Me.rowItemTipeKontrak)
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttanchmentKontrak_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DGVAttanchmentKontrak.UserDeletingRow
        Dim identifier As String
        Try
            identifier = Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, Me.DGVKontrak.CurrentRow.Index).Value
            Me.tblItemAttch.DefaultView.RowFilter = "[identifier kontrak] = '" & identifier & "' AND [desc] = '" & Me.DGVAttanchmentKontrak.Item(Me.ColKontrakAtt_Description.Index, Me.DGVAttanchmentKontrak.CurrentRow.Index).Value & "'"
            For x As Integer = 0 To Me.tblItemAttch.DefaultView.Count - 1
                If Me.tblItemAttch.DefaultView.Count > 0 Then
                    Me.rowItemAttch = Me.tblItemAttch.DefaultView(0).Row
                    Me.tblItemAttch.Rows.Remove(Me.rowItemAttch)
                End If
            Next
           
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttanchment_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVAttanchment.Sorted
        Try
            Me.ObjCekdata.SetNumber(Me.DGVAttanchment)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVAttanchment_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DGVAttanchment.UserDeletedRow
        Try
            Me.ObjCekdata.SetNumber(Me.DGVAttanchment)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    
    Private Sub TabControl2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl2.SelectedIndexChanged
        Dim ObjUser As New ClsUser

        Try
            'pastikan user memiliki hak akses untuk melihat detail pembelian 
            If Me.TabControl2.SelectedIndex = 1 Then
                If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmMasterKaryawan_LihatKompetensi") = False Then
                    Me.TabControl2.SelectedIndex = 0
                    MsgBox("You don't have authorization form this information.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            'If Me.DGVKontrak.Rows.Count <> 0 And Me.TabControl2.SelectedIndex = 2 Then
            '    reloadDataTableKontrak(Me.DGVKontrak.Item(Me.ColContract_IdentifierContract.Index, 0).Value)
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjUser = Nothing
        End Try
    End Sub

    Private Sub chkUseOvertimeOrder_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkUseOvertimeOrder.CheckedChanged
        Try
            If Me.chkUseOvertimeOrder.Checked = True Then
                Me.chkSendNotifEmailOvertimeOrder.Enabled = True
            ElseIf Me.chkUseOvertimeOrder.Checked = False Then
                Me.chkSendNotifEmailOvertimeOrder.Enabled = False
                Me.chkSendNotifEmailOvertimeOrder.Checked = False
            End If
            'If Me.chkUseOvertimeOrder.Checked = True And Me.chkUseOvertimeOrder.Enabled = True Then
            '    Me.chkSendNotifEmailOvertimeOrder.Enabled = True
            'ElseIf Me.chkUseOvertimeOrder.Checked = True And Me.chkUseOvertimeOrder.Enabled = False Then
            '    Me.chkSendNotifEmailOvertimeOrder.Enabled = False
            'ElseIf Me.chkUseOvertimeOrder.Checked = False And Me.chkSendNotifEmailOvertimeOrder.Enabled = True Then
            '    Me.chkSendNotifEmailOvertimeOrder.Enabled = False
            '    Me.chkSendNotifEmailOvertimeOrder.Checked = False
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetUseOvertimeOrder()
        Try
            If Me.ChkOverrideOT.Checked = True And Me.ChkOverrideEarlyOT.Checked = True Then
                If Me.CboIsCalculateEarlyOT.SelectedIndex = 1 And Me.CboIsCalculateOT.SelectedIndex = 1 Then
                    'Me.chkUseOvertimeOrder.Enabled = False
                    'Me.chkSendNotifEmailOvertimeOrder.Enabled = False
                    'Me.chkUseOvertimeOrder.Checked = False
                    'Me.chkSendNotifEmailOvertimeOrder.Checked = False
                    Me.ChkOverrideOTOrder.Enabled = False
                    Me.ChkOverrideOTOrder.Checked = False
                Else
                    'Me.chkUseOvertimeOrder.Enabled = True
                    'Me.chkSendNotifEmailOvertimeOrder.Enabled = True
                    Me.ChkOverrideOTOrder.Enabled = True
                End If
            Else
                Me.ChkOverrideOTOrder.Enabled = True
                If Me.ChkOverrideOTOrder.Checked = True Then
                    Me.chkUseOvertimeOrder.Enabled = True
                End If
                'Me.chkUseOvertimeOrder.Enabled = True
                'Me.chkSendNotifEmailOvertimeOrder.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkOverrideOTOrder_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkOverrideOTOrder.CheckedChanged
        Try
            If Me.ChkOverrideOTOrder.Checked = True Then
                Me.chkUseOvertimeOrder.Enabled = True
            Else
                Me.chkUseOvertimeOrder.Enabled = False
                Me.chkUseOvertimeOrder.Checked = False
                Me.chkSendNotifEmailOvertimeOrder.Enabled = False
                Me.chkSendNotifEmailOvertimeOrder.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbCopy.Click
        Dim hj As New FrmMasterKaryawanCari

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryEmploye = -1 Then
                Me.KosongkanField()

                If Me.ObjKaryawan.CariDetail(hj.PrimaryEmploye, Me, Me.dgvFamily, Me.DGVKompetensi) = True Then
                    Me.SalinKeField()

                    Me.Baru = True
                    Me.Aktifkandetail()
                    Me.TxtEmployeeCode.Focus()
                    Me.TabControl1.SelectedIndex = 0
                    Me.TabControl2.SelectedIndex = 0
                    Me.TxtPrimaryEmployee.Text = -1

                    Me.TxtEmployeeCode.Text = "COPY OF " + Me.TxtEmployeeCode.Text
                    If Me.TxtNIK.Text <> "" Then
                        Me.TxtNIK.Text = "COPY OF " + Me.TxtNIK.Text
                    End If

                    Me.ToolStripStatusLabel1.Text = "Created By: " + Me.UserCode
                    Me.ToolStripStatusLabel2.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                    Me.ToolStripStatusLabel3.Text = "Modified By: " + Me.UserCode
                    Me.ToolStripStatusLabel4.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class