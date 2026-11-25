Option Explicit On
Option Strict Off

Public Class clsMasterSupplier
    Dim ObjCD As New Core_Module.ClsCekData

    Private mprimary As Double = -1
    Private mprimaryDefaultCurrency As Double = -1
    Private mSupplierCode As String = ""
    Private mSupplierName As String = ""
    Private mCountry As String = ""
    Private mPaymentTerm As String = ""
    Private mLeadTime As Integer = 0
    Private mNPWP As String = ""
    Private mDefaulPriceType As EnumPriceType = EnumPriceType.Exclude
    Private mCompanyWebsite As String = ""
    Private mForeignSupplier As Boolean = False

    'Comment: Perubahan NPWP No 16 Digit
    Private mNPWPLama As String = ""
    Private mNPWPType As ClsTrans_SettingNPWPNo.TipeNPWP = ClsTrans_SettingNPWPNo.TipeNPWP.Individual

    'Comment: Perubahan PPN 12%
    Private mTKU_ID As String = ""
    Private mEmail As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Enum EnumDefaultFindBy As Integer
        ByCode
        ByDesc
    End Enum

    'Comment: Perubahan Perubahan PPN 12%
    Public Property TKU_ID() As String
        Get
            Return Me.mTKU_ID
        End Get
        Set(ByVal value As String)
            mTKU_ID = value
        End Set
    End Property

    Public Property Email() As String
        Get
            Return Me.mEmail
        End Get
        Set(ByVal value As String)
            mEmail = value
        End Set
    End Property

    'Tambahan NPWP
    Public Property NPWPType() As ClsTrans_SettingNPWPNo.TipeNPWP
        Get
            Return Me.mNPWPType
        End Get
        Set(ByVal value As ClsTrans_SettingNPWPNo.TipeNPWP)
            mNPWPType = value
        End Set
    End Property

    Public Property NPWPLama() As String
        Get
            Return Me.mNPWPLama
        End Get
        Set(ByVal value As String)
            mNPWPLama = value
        End Set
    End Property
    'END Tambahan NPWP

    Public Enum EnumDefaultFindBy_SettingUser As Integer
        ByGlobalSetting
        ByCode
        ByDesc
    End Enum

    Public Enum EnumPriceType As Integer
        Exclude
        Include
    End Enum

    Public Property ForeignSupplier() As Boolean
        Get
            Return Me.mForeignSupplier
        End Get
        Set(ByVal value As Boolean)
            mForeignSupplier = value
        End Set
    End Property

    Public Property CompanyWebsite() As String
        Get
            Return Me.mCompanyWebsite
        End Get
        Set(ByVal value As String)
            mCompanyWebsite = value
        End Set
    End Property

    Public Property PriceType() As EnumPriceType
        Get
            Return Me.mDefaulPriceType
        End Get
        Set(ByVal value As EnumPriceType)
            mDefaulPriceType = value
        End Set
    End Property

    Public Property LeadTime() As Integer
        Get
            Return Me.mLeadTime
        End Get
        Set(ByVal value As Integer)
            mLeadTime = value
        End Set
    End Property

    Public Property PrimaryCurrency() As Double
        Get
            Return Me.mprimaryDefaultCurrency
        End Get
        Set(ByVal value As Double)
            mprimaryDefaultCurrency = value
        End Set
    End Property

    Public Property Primary() As Double
        Get
            Return mprimary
        End Get
        Set(ByVal value As Double)
            mprimary = value
        End Set
    End Property

    Public Property SupplierCode() As String
        Get
            Return Me.mSupplierCode
        End Get
        Set(ByVal value As String)
            mSupplierCode = value
        End Set
    End Property

    Public Property SupplierName() As String
        Get
            Return Me.mSupplierName
        End Get
        Set(ByVal value As String)
            mSupplierName = value
        End Set
    End Property

    Public Property country() As String
        Get
            Return mCountry
        End Get
        Set(ByVal value As String)
            mCountry = value
        End Set
    End Property

    Public Property PaymentTerm() As String
        Get
            Return mPaymentTerm
        End Get
        Set(ByVal value As String)
            mPaymentTerm = value
        End Set
    End Property

    Public Property NPWP() As String
        Get
            Return mNPWP
        End Get
        Set(ByVal value As String)
            mNPWP = value
        End Set
    End Property

    Public Property CreatedBy() As String
        Get
            Return mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property ModifiedBy() As String
        Get
            Return mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Public Property ModifiedDate() As Date
        Get
            Return mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property TimeStamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
        End Set
    End Property

    Public Function SimpanBaru(ByVal vusercode As String, ByVal dgvCP As DataGridView, ByVal dgvAdd As DataGridView, _
    ByVal dgvBank As DataGridView) As Boolean
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter
        Dim DTA1 As New dseMasterSupplierTableAdapters.tblmaster_supplier_itemcpTableAdapter
        Dim DTA2 As New dseMasterSupplierTableAdapters.tblmaster_supplier_item_addressTableAdapter
        Dim DTA3 As New dseMasterSupplierTableAdapters.tblmaster_supplier_item_bankTableAdapter
        Dim row As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString

            If DTA.Insert(Me.mSupplierCode, mSupplierName, Me.mDefaulPriceType, Me.mPaymentTerm, Me.mLeadTime, _
            mCountry, Me.mprimaryDefaultCurrency, Me.mForeignSupplier, Me.mNPWP, Me.mCompanyWebsite, _
            vusercode, Now, vusercode, Now, Now, Me.mNPWPLama, Me.mNPWPType, Me.mTKU_ID, Me.mEmail) > 0 Then
                Me.mprimary = DTA.ScalarQuery_GetPrimary(Me.mSupplierCode).Value
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mprimary).Value
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vusercode
                Me.mModifiedBy = vusercode

                'simpan data contact person
                For Each row In dgvCP.Rows
                    With row
                        If row.Index <> dgvCP.Rows.Count - 1 Then
                            DTA1.Insert(Me.mprimary, .Cells(1).Value, .Cells(2).Value, .Cells(3).Value, _
                            .Cells(4).Value, .Cells(5).Value, .Cells(6).Value, CStr(Me.mprimary) + "-" + CStr(row.Index))
                        End If
                    End With
                Next
                'simpan data address
                For Each row In dgvAdd.Rows
                    With row
                        If row.Index <> dgvAdd.Rows.Count - 1 Then
                            DTA2.Insert(Me.mprimary, .Cells(1).Value, .Cells(2).Value, .Cells(3).Value, _
                             .Cells(4).Value, .Cells(5).Value, .Cells(6).Value, .Cells(7).Value, .Cells(8).Value, _
                             .Cells(9).Value, CStr(Me.mprimary) + "-" + CStr(.Index))
                        End If
                    End With
                Next
                'simpan data bank
                For Each row In dgvBank.Rows
                    With row
                        If row.Index <> dgvBank.Rows.Count - 1 Then
                            DTA3.Insert(Me.mprimary, .Cells(1).Value, .Cells(2).Value, .Cells(8).Value, _
                             .Cells(4).Value, .Cells(5).Value, .Cells(6).Value, .Cells(7).Value, CStr(Me.mprimary) _
                             + "-" + CStr(row.Index))
                        End If
                    End With
                Next

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapusAddress(ByVal IdentifierAddress As String) As Boolean
        Dim DTAPO As New dseMasterSupplierTableAdapters.tbltrans_poTableAdapter

        Try
            DTAPO.Connection.ConnectionString = Core_Module.GetConstring
            'tidak boleh hapus bila:
            '1. sudah digunakan pada PO
            If DTAPO.ScalarQuery_GetJml(IdentifierAddress).HasValue = True Then
                If DTAPO.ScalarQuery_GetJml(IdentifierAddress).Value > 0 Then
                    MsgBox("There purchase order refering to this address, you cannot delete this data")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPO.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapusItemCP(ByVal IdentifierCp As String) As Boolean
        Try
            'tidak boleh hapus bila:
            'msgbox("belum")
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehHapusItemBank(ByVal IdentifierBank As String) As Boolean
        Try
            'tidak boleh hapus bila:
            'msgbox("belum")
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vusercode As String, ByVal dgvCP As DataGridView, ByVal dgvAdd As DataGridView, _
    ByVal dgvBank As DataGridView) As Boolean
        Dim dta As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter
        Dim DTA1 As New dseMasterSupplierTableAdapters.tblmaster_supplier_itemcpTableAdapter
        Dim DTA2 As New dseMasterSupplierTableAdapters.tblmaster_supplier_item_addressTableAdapter
        Dim DTA3 As New dseMasterSupplierTableAdapters.tblmaster_supplier_item_bankTableAdapter
        Dim row As DataGridViewRow

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = dta.Connection.ConnectionString
            DTA2.Connection.ConnectionString = dta.Connection.ConnectionString
            DTA3.Connection.ConnectionString = dta.Connection.ConnectionString

            If dta.Update(Me.mSupplierCode, mSupplierName, Me.mDefaulPriceType, Me.mPaymentTerm, Me.mLeadTime, _
                Me.mCountry, Me.mprimaryDefaultCurrency, Me.mForeignSupplier, Me.mNPWP, Me.mCompanyWebsite, _
                vusercode, Now, Now, Me.mNPWPLama, Me.mNPWPType, Me.mTKU_ID, Me.mEmail, mprimary, mTimestamp) > 0 Then
                Me.mTimestamp = dta.ScalarQuery_GetTimestamp(Me.mprimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vusercode

                'hapus data item lama
                DTA1.Delete(Me.mprimary)
                DTA2.Delete(Me.mprimary)
                DTA3.Delete(Me.mprimary)
                'simpan data item baru
                'simpan data contact person
                For Each row In dgvCP.Rows
                    With row
                        If row.Index <> dgvCP.Rows.Count - 1 Then
                            DTA1.Insert(Me.mprimary, .Cells(1).Value, .Cells(2).Value, .Cells(3).Value, _
                            .Cells(4).Value, .Cells(5).Value, .Cells(6).Value, CStr(Me.mprimary) + "-" + CStr(row.Index))
                        End If
                    End With
                Next
                'simpan data address
                For Each row In dgvAdd.Rows
                    With row
                        If row.Index <> dgvAdd.Rows.Count - 1 Then
                            DTA2.Insert(Me.mprimary, .Cells(1).Value, .Cells(2).Value, .Cells(3).Value, _
                             .Cells(4).Value, .Cells(5).Value, .Cells(6).Value, .Cells(7).Value, .Cells(8).Value, _
                             .Cells(9).Value, CStr(Me.mprimary) + "-" + CStr(.Index))
                        End If
                    End With
                Next
                'simpan data bank
                For Each row In dgvBank.Rows
                    With row
                        If row.Index <> dgvBank.Rows.Count - 1 Then
                            DTA3.Insert(Me.mprimary, .Cells(1).Value, .Cells(2).Value, .Cells(8).Value, _
                             .Cells(4).Value, .Cells(5).Value, .Cells(6).Value, .Cells(7).Value, CStr(Me.mprimary) _
                             + "-" + CStr(row.Index))
                        End If
                    End With
                Next

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            DTA1.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vprimary As Double) As Boolean
        Dim dta As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vprimary) = True Then
                dta.Delete(vprimary)    'secara otomatis data item juga akan dihapus
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function ApakahLokal(ByVal vPrimarySupplier As Double) As Boolean
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_IsForeignSupplier(vPrimarySupplier).HasValue = True Then
                If DTA.ScalarQuery_IsForeignSupplier(vPrimarySupplier).Value > 0 Then
                    Return False
                Else
                    Return True
                End If
            Else
                'nilai balikan default adalah supplier lokal
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vSupplierCode As String) As Boolean
        Dim Dta As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Dta.ScalarQuery_GetJml(vSupplierCode).HasValue Then
                If Dta.ScalarQuery_GetJml(vSupplierCode).Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vprimary As Double, Optional ByVal vTampilkanPesan As Boolean = True) As Boolean
        Dim DTAPO As New dseMasterSupplierTableAdapters.tbltrans_poTableAdapter
        Dim DTA_IPO As New dseMasterSupplierTableAdapters.tbltrans_ipoTableAdapter
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter
        Dim Jml As Nullable(Of Long)
        Dim TipeAOPDP As Integer = 1
        Dim TipeTransaksiCN As Integer = 0
        Dim TipeTransaksiDN As Integer = 1
        Dim TipeKartuCNDN As Integer = 1
        Dim TipeQuickPR As Integer = 1

        Try
            DTAPO.Connection.ConnectionString = Core_Module.GetConstring
            DTA_IPO.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'jika sudah pernah digunakan pada PO maka tidak bisa dihapus
            If DTAPO.ScalarQuery_GetJmlByPO(vprimary).HasValue = True Then
                If DTAPO.ScalarQuery_GetJmlByPO(vprimary).Value > 0 Then
                    If vTampilkanPesan = True Then
                        MsgBox("There are purchase orders refering to this supplier. You cannot delete this data", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            'credit - note
            Jml = DTA.ScalarQuery_GetJmlCNDN(vprimary, TipeTransaksiCN, TipeKartuCNDN)
            If Jml.HasValue = True Then
                If Jml.Value > 0 Then
                    If vTampilkanPesan = True Then
                        MsgBox("There are credit notes refering to this supplier. You cannot delete this data", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            'debit - note 
            Jml = DTA.ScalarQuery_GetJmlCNDN(vprimary, TipeTransaksiDN, TipeKartuCNDN)
            If Jml.HasValue = True Then
                If Jml.Value > 0 Then
                    If vTampilkanPesan = True Then
                        MsgBox("There are debit notes refering to this supplier. You cannot delete this data", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            If DTA_IPO.ScalarQuery_GetJmlByPrimarySupplier(vprimary).HasValue = True Then
                If DTA_IPO.ScalarQuery_GetJmlByPrimarySupplier(vprimary).Value > 0 Then
                    If vTampilkanPesan = True Then
                        MsgBox("There are internal purchase orders refering to this supplier. You cannot delete this data", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            'aop dp
            Jml = DTA.ScalarQuery_GetJmlAOPDP(TipeAOPDP, vprimary)
            If Jml.HasValue = True Then
                If Jml.Value > 0 Then
                    If vTampilkanPesan = True Then
                        MsgBox("There are down payment AOP refering to this supplier. You cannot delete this data", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            'quick pr
            Jml = DTA.ScalarQuery_GetJmlQuickPR(TipeQuickPR, vprimary)
            If Jml.HasValue = True Then
                If Jml.Value > 0 Then
                    If vTampilkanPesan = True Then
                        MsgBox("There are quick purchase return refering to this supplier. You cannot delete this data", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            'begining balance purchase  
            If DTAPO.ScalarQuery_GetJmlBB(vprimary).HasValue = True Then
                If DTAPO.ScalarQuery_GetJmlBB(vprimary).Value > 0 Then
                    If vTampilkanPesan = True Then
                        MsgBox("There are purchase begining balance refering to this supplier. You cannot delete this data", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            'begining balance purchase return
            If DTAPO.ScalarQuery_GetJmlBBReturn(vprimary).HasValue = True Then
                If DTAPO.ScalarQuery_GetJmlBBReturn(vprimary).Value > 0 Then
                    If vTampilkanPesan = True Then
                        MsgBox("There are purchase return begining balance refering to this supplier. You cannot delete this data", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPO.Dispose()
            DTA_IPO.Dispose()
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPaymentTerm(ByVal vPrimarySupplier As Double) As Double
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter
        Dim RetvalSTR As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetPaymentTerm(vPrimarySupplier)) = True Then
                Return 0
            Else
                RetvalSTR = DTA.ScalarQuery_GetPaymentTerm(vPrimarySupplier)
                If IsNumeric(RetvalSTR) = False Then
                    Return 0
                Else
                    Return CDbl(RetvalSTR)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal SupplierCode As String) As Double
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(SupplierCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(SupplierCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSupplierName(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetSupplierName(vPrimary)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetSupplierName(vPrimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetKodeAddress(ByVal vIdentifierAdd As String) As String
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplier_item_address1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_getAddressCode(vIdentifierAdd)) = False Then
                Return DTA.ScalarQuery_getAddressCode(vIdentifierAdd)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNPWP(ByVal vPrimarySuppl As Double) As String
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoNPWP(vPrimarySuppl)) = False Then
                Return DTA.ScalarQuery_GetNoNPWP(vPrimarySuppl)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPhoneNo(ByVal vPrimarySuppl As Double, ByVal vAddIdentifier As String) As String
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplier_item_address1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_getPhoneNo(vPrimarySuppl, vAddIdentifier)) = False Then
                Return DTA.ScalarQuery_getPhoneNo(vPrimarySuppl, vAddIdentifier)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiTxtAddress(ByVal vIdentifierAdd As String, ByVal vPrimarySupplier As Double, ByVal txt As TextBox) As Boolean
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplier_item_address1TableAdapter
        Dim TBL As New dseMasterSupplier.tblmaster_supplier_item_address1DataTable
        Dim row As dseMasterSupplier.tblmaster_supplier_item_address1Row
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_PrimarySupNIdentifier(TBL, vPrimarySupplier, vIdentifierAdd) > 0 Then
                row = TBL.Rows(0)
                txt.Clear()
                txt.Text = row.Line_1
                txt.Text = txt.Text + vbCrLf + row.Line_2
                txt.Text = txt.Text + vbCrLf + row.Line_3
                txt.Text = txt.Text + vbCrLf + row.Line_4
                txt.Text = txt.Text + vbCrLf + row.Line_5
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetIdentifierAdd(ByVal vPrimarySupplier As Double) As String
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplier_item_addressTableAdapter
        Dim TBL As New dseMasterSupplier.tblmaster_supplier_item_addressDataTable
        Dim row As dseMasterSupplier.tblmaster_supplier_item_addressRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimarySupplier2(TBL, vPrimarySupplier)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Return row.Identifier
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetIdentifierAdd(ByVal vPrimarySupl As Double, ByVal vKodeAdd As String) As String
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplier_item_address1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_getIdentifier(vPrimarySupl, vKodeAdd)) = False Then
                Return DTA.ScalarQuery_getIdentifier(vPrimarySupl, vKodeAdd)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSupplierCode(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCode(vPrimary)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetCode(vPrimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlBy_PrimaryCurrency(ByVal vPrimaryCurrency As Double) As Double
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter
        Dim DTA1 As New dseMasterSupplierTableAdapters.tblmaster_supplier_item_bankTableAdapter
        Dim Dummy As Double = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            If DTA.ScalarQuery_GetJmlByCurrency(vPrimaryCurrency).HasValue = True Then
                Dummy += DTA.ScalarQuery_GetJmlByCurrency(vPrimaryCurrency).Value
            End If
            If DTA1.ScalarQuery_GetJml(vPrimaryCurrency).HasValue = True Then
                Dummy += DTA1.ScalarQuery_GetJml(vPrimaryCurrency).Value
            End If

            Return Dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKodeSupplier As String) As Double
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetJml(vKodeSupplier)) = False Then
                Return DTA.ScalarQuery_GetJml(vKodeSupplier)
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
           EnumSearchCriteria) As Double
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplier1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                If DTA.ScalarQuery_GetJml(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                If DTA.ScalarQuery_GetJml2(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml2(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByAccountNo Then
                If DTA.ScalarQuery_GetJml3(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml3(vKriteria + "%").Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
     ByVal DGV As DataGridView, ByVal TBL As dseMasterSupplier.tblmaster_supplier1DataTable, _
     ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplier1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByAccountNo Then
                QueryResult = DTA.FillBy_AccountNo(TBL, Criteria + "%")
            End If

            '2025 Core Tax Tahap 2.1, ganti npwp ke format 16
            Dim row As dseMasterSupplier.tblmaster_supplier1Row
            Dim objSetNPWP As New ClsTrans_SettingNPWPNo
            If Today.Date >= objSetNPWP.Tglberlaku Then
                For Each row In TBL.Rows
                    row.NPWP_No = objSetNPWP.format16DigitSpasi(row.NPWP_No)
                Next
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Supplier code"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
            End If


            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetTimestamp(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal dgvAddress As DataGridView, _
    ByVal dgvItemCP As DataGridView, ByVal dgvItemBank As DataGridView) As Boolean
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter
        Dim TBL As New dseMasterSupplier.tblmaster_supplierDataTable
        Dim row As dseMasterSupplier.tblmaster_supplierRow
        Dim DTA1 As New dseMasterSupplierTableAdapters.tblmaster_supplier_item_addressTableAdapter
        Dim TBL1 As New dseMasterSupplier.tblmaster_supplier_item_addressDataTable
        Dim row1 As dseMasterSupplier.tblmaster_supplier_item_addressRow
        Dim DTA2 As New dseMasterSupplierTableAdapters.tblmaster_supplier_itemcpTableAdapter
        Dim TBL2 As New dseMasterSupplier.tblmaster_supplier_itemcpDataTable
        Dim row2 As dseMasterSupplier.tblmaster_supplier_itemcpRow
        Dim DTA3 As New dseMasterSupplierTableAdapters.tblmaster_supplier_item_bankTableAdapter
        Dim TBL3 As New dseMasterSupplier.tblmaster_supplier_item_bankDataTable
        Dim row3 As dseMasterSupplier.tblmaster_supplier_item_bankRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.FillBy_Primary(TBL, vPrimary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mCompanyWebsite = .Company_Website
                    Me.mCountry = .Country
                    Me.mDefaulPriceType = .Price_Type
                    Me.mForeignSupplier = .Foreign_Supplier
                    Me.mLeadTime = .Lead_Time
                    Me.mNPWP = .NPWP_No
                    Me.mPaymentTerm = .Payment_Term
                    Me.mprimary = .Primary
                    Me.mprimaryDefaultCurrency = .Primary_Default_Currency
                    Me.mSupplierCode = .Supplier_Code
                    Me.mSupplierName = .Supplier_Name

                    'salin data addres
                    dgvAddress.Rows.Clear()
                    dgvItemBank.Rows.Clear()
                    dgvItemCP.Rows.Clear()

                    DTA1.FillBy_PrimarySupplier(TBL1, Me.mprimary)
                    Dim Obj2 As New ClsCekData
                    For Each row1 In TBL1.Rows
                        With row1
                            dgvAddress.Rows.Add("-", .Address_Code, .Line_1, .Line_2, .Line_3, .Line_4 _
                            , .Line_5, .Phone_No, .Fax_No, .Remark, .Identifier)
                        End With
                    Next
                    'salin data contact person
                    DTA2.FillBy_PrimarySupplier(TBL2, Me.mprimary)
                    For Each row2 In TBL2.Rows
                        With row2
                            dgvItemCP.Rows.Add("-", .Name, .Position, .Phone_No, .Fax_No, _
                            ._E_mail_Address, .Remark, .Identifier)
                        End With
                    Next
                    'salin data bank
                    Dim ObjCurrency As New clsMasterCurrency
                    DTA3.FillBy_PrimarySupplier(TBL3, Me.mprimary)
                    For Each row3 In TBL3.Rows
                        With row3
                            dgvItemBank.Rows.Add("-", ._Bank_s_Name, .Account_No, ObjCurrency.GetCurrencyCode(.Primary_Currency), _
                            .Address, .Beneficiary_Name, .Phone_No, .Remark, .Primary_Currency, .Identifier)
                        End With
                    Next

                    Obj2.SetNumber(dgvAddress)
                    Obj2.SetNumber(dgvItemCP)
                    Obj2.SetNumber(dgvItemBank)
                    Obj2 = Nothing
                    ObjCurrency = Nothing
                    'update NPWP 16 Digit
                    Me.mNPWPLama = .NPWP_Old
                    Me.mNPWPType = .NPWP_Type

                    'update PPN 12%
                    Me.mTKU_ID = .TKU_ID
                    Me.mEmail = .Email

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
            DTA2.Dispose()
            TBL2.Dispose()
            DTA3.Dispose()
            TBL3.Dispose()
        End Try
    End Function

    Public Function IsiCheckListBox(ByVal lbo As CheckedListBox, ByVal TBL As dseMasterSupplier.tblmaster_supplier2DataTable) As Boolean
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplier2TableAdapter
        Dim row As dseMasterSupplier.tblmaster_supplier2Row
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            lbo.DataSource = TBL.DefaultView
            For Each row In TBL.Rows
                row.Combination = row.Supplier_Code + " - " + row.Supplier_Name
            Next
            lbo.DisplayMember = "Combination"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiCboAddress(ByVal cbo As ComboBox, ByVal PrimarySupplier As Double, _
    ByVal TBL As dseMasterSupplier.tblmaster_supplier_item_addressDataTable) As Boolean
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplier_item_addressTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            cbo.DataSource = Nothing
            DTA.FillBy_PrimarySupplier(TBL, PrimarySupplier)
            cbo.DataSource = TBL.DefaultView
            TBL.DefaultView.Sort = "[Address Code]"
            cbo.DisplayMember = "Address Code"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiCboAddress(ByVal cbo As ComboBox, ByVal PrimarySupplier As Double) As Boolean
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplier_item_addressTableAdapter
        Dim TBL As New dseMasterSupplier.tblmaster_supplier_item_addressDataTable
        Dim row As dseMasterSupplier.tblmaster_supplier_item_addressRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            cbo.Items.Clear()
            If DTA.FillBy_PrimarySupplier(TBL, PrimarySupplier) > 0 Then
                For Each row In TBL.Rows
                    cbo.Items.Add(row.Address_Code)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetPrimaryDefaultCurrency(ByVal vPrimarySupplier As Double) As Double
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetDefaultSupplierCurrency(vPrimarySupplier).HasValue = False Then
                Return -1
            Else
                Return DTA.ScalarQuery_GetDefaultSupplierCurrency(vPrimarySupplier).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub Get_SuppName_ContactPerson_PhoneNo(ByVal vPrimarySupplier As Double, ByRef vSupplierName As String, _
    ByRef vContactPerson As String, ByRef vPhoneNo As String)
        Dim DTA_CP As New dseMasterSupplierTableAdapters.tblmaster_supplier_itemcpTableAdapter
        Dim TBL_CP As New dseMasterSupplier.tblmaster_supplier_itemcpDataTable
        Dim row As dseMasterSupplier.tblmaster_supplier_itemcpRow

        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA_CP.Connection.ConnectionString = DTA.Connection.ConnectionString

            'nama
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetSupplierName(vPrimarySupplier)) = False Then
                vSupplierName = DTA.ScalarQuery_GetSupplierName(vPrimarySupplier)
            Else
                vSupplierName = ""
            End If

            'contact person dan no telepon
            vContactPerson = ""
            vPhoneNo = ""

            DTA_CP.FillBy_PrimarySupplier(TBL_CP, vPrimarySupplier)
            For Each row In TBL_CP.Rows
                If vContactPerson = "" Then
                    vContactPerson = row.Name
                Else
                    vContactPerson = vContactPerson + ", " + row.Name
                End If
                If vPhoneNo = "" Then
                    vPhoneNo = row.Phone_No
                Else
                    vPhoneNo = vPhoneNo + ", " + row.Phone_No
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_CP.Dispose()
            TBL_CP.Dispose()
            DTA.Dispose()
        End Try
    End Sub

    Public Function Get_ContactPerson(ByVal vPrimarySupplier As Double) As String
        Dim DTA_CP As New dseMasterSupplierTableAdapters.tblmaster_supplier_itemcpTableAdapter
        Dim TBL_CP As New dseMasterSupplier.tblmaster_supplier_itemcpDataTable
        Dim row As dseMasterSupplier.tblmaster_supplier_itemcpRow
        Dim RetVal As String = ""

        Try
            DTA_CP.Connection.ConnectionString = Core_Module.GetConstring

            DTA_CP.FillBy_PrimarySupplier(TBL_CP, vPrimarySupplier)
            For Each row In TBL_CP.Rows
                If RetVal = "" Then
                    RetVal = row.Name
                Else
                    RetVal = RetVal + ", " + row.Name
                End If
            Next
            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA_CP.Dispose()
            TBL_CP.Dispose()
        End Try
    End Function


    Public Function GetNPWPOld(ByVal vPrimarySuppl As Double) As String
        Dim DTA As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoNPWPOld(vPrimarySuppl)) = False Then
                Return DTA.ScalarQuery_GetNoNPWPOld(vPrimarySuppl)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
