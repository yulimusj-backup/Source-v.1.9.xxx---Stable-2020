Option Explicit On
Option Strict Off

Public Class clsMasterCustomer
    Dim ObjCD As New Core_Module.ClsCekData

    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianMaster
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Private mprimary As Double = -1
    Private mprimaryCustomerGroup As Double = -1
    Private mprimarySalesArea As Double = -1
    Private mprimaryCustomerSegmentation As Double = -1
    Private mprimaryCustomerType As Double = -1

    Private mCustomerCode As String = ""
    Private mcustomerName As String = ""
    Private mCity As String = ""
    Private mCountry As String = ""
    Private mZipCode As String = ""

    Private mPaymentTerm As Double = -1
    Private mCreditLimit As Double = -1
    Private mNPWP As String = ""
    Private mNPWP_Name As String = ""
    Private mDefaultTaxCode As String = ""
    Private mTaxAddressLine1 As String = ""
    Private mTaxAddressLine2 As String = ""
    Private mRemarks As String = ""
    Private mPrimarySalesPerson As Double = -1

    Private mEFak_NamaJalan As String = ""
    Private mEFak_Blok As String = ""
    Private mEFak_No As String = ""
    Private mEFak_RT As String = ""
    Private mEFak_RW As String = ""
    Private mEFak_Kecamatan As String = ""
    Private mEFak_Kelurahan As String = ""
    Private mEFak_Kabupaten As String = ""
    Private mEFak_Propinsi As String = ""
    Private mEFak_KodePos As String = ""
    Private mEFak_Telepon As String = ""
    '//TR : Tambahan baru Nik/Pas
    Private mOptionNIKPASS As Integer = 0
    Private mNoNIKPASS As String = ""
    Private mPrimaryCustomerReference As Double = -1
    Private mJoinDate As Date = Now

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

    Public Enum EnumDefaultFindBy_SettingUser As Integer
        ByGlobalSetting
        ByCode
        ByDesc
    End Enum
    '//TR : Tambahan baru Nik/Pas
    Public Enum EnumOptionNikPass As Integer
        ByNIK
        ByPASSPORT
        ByOtherID '2025 Core Tax Tahap 2.1
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

    Public Property OptionNIKPASS() As Integer
        Get
            Return Me.mOptionNIKPASS
        End Get
        Set(ByVal value As Integer)
            mOptionNIKPASS = value
        End Set
    End Property

    Public Property NoNIKPASS() As String
        Get
            Return Me.mNoNIKPASS
        End Get
        Set(ByVal value As String)
            mNoNIKPASS = value
        End Set
    End Property

    Public Property EFak_Telepon() As String
        Get
            Return Me.mEFak_Telepon
        End Get
        Set(ByVal value As String)
            mEFak_Telepon = value
        End Set
    End Property

    Public Property EFak_KodePos() As String
        Get
            Return Me.mEFak_KodePos
        End Get
        Set(ByVal value As String)
            mEFak_KodePos = value
        End Set
    End Property

    Public Property EFak_Propinsi() As String
        Get
            Return Me.mEFak_Propinsi
        End Get
        Set(ByVal value As String)
            mEFak_Propinsi = value
        End Set
    End Property

    Public Property EFak_Kabupaten() As String
        Get
            Return Me.mEFak_Kabupaten
        End Get
        Set(ByVal value As String)
            mEFak_Kabupaten = value
        End Set
    End Property

    Public Property EFak_Kelurahan() As String
        Get
            Return Me.mEFak_Kelurahan
        End Get
        Set(ByVal value As String)
            mEFak_Kelurahan = value
        End Set
    End Property

    Public Property EFak_Kecamatan() As String
        Get
            Return Me.mEFak_Kecamatan
        End Get
        Set(ByVal value As String)
            mEFak_Kecamatan = value
        End Set
    End Property

    Public Property EFak_RW() As String
        Get
            Return Me.mEFak_RW
        End Get
        Set(ByVal value As String)
            mEFak_RW = value
        End Set
    End Property

    Public Property EFak_RT() As String
        Get
            Return Me.mEFak_RT
        End Get
        Set(ByVal value As String)
            mEFak_RT = value
        End Set
    End Property

    Public Property EFak_No() As String
        Get
            Return Me.mEFak_No
        End Get
        Set(ByVal value As String)
            mEFak_No = value
        End Set
    End Property

    Public Property EFak_Blok() As String
        Get
            Return Me.mEFak_Blok
        End Get
        Set(ByVal value As String)
            mEFak_Blok = value
        End Set
    End Property

    Public Property EFak_NamaJalan() As String
        Get
            Return Me.mEFak_NamaJalan
        End Get
        Set(ByVal value As String)
            mEFak_NamaJalan = value
        End Set
    End Property

    Public Property PrimarySalesPerson() As Double
        Get
            Return Me.mPrimarySalesPerson
        End Get
        Set(ByVal value As Double)
            mPrimarySalesPerson = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return Me.mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
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

    Public Property CustomerCode() As String
        Get
            Return Me.mCustomerCode
        End Get
        Set(ByVal value As String)
            mCustomerCode = value
        End Set
    End Property

    Public Property CustomerName() As String
        Get
            Return Me.mcustomerName
        End Get
        Set(ByVal value As String)
            mcustomerName = value
        End Set
    End Property

    Public Property PrimarysSalesArea() As Double
        Get
            Return mprimarySalesArea
        End Get
        Set(ByVal value As Double)
            mprimarySalesArea = value
        End Set
    End Property

    Public Property PrimaryCustomerSegmentation() As Double
        Get
            Return mprimaryCustomerSegmentation
        End Get
        Set(ByVal value As Double)
            mprimaryCustomerSegmentation = value
        End Set
    End Property

    Public Property PrimaryCustomertype() As Double
        Get
            Return mprimaryCustomerType
        End Get
        Set(ByVal value As Double)
            mprimaryCustomerType = value
        End Set
    End Property

    Public Property PrimaryCustomerGroup() As Double
        Get
            Return mprimaryCustomerGroup
        End Get
        Set(ByVal value As Double)
            mprimaryCustomerGroup = value
        End Set
    End Property

    Public Property city() As String
        Get
            Return mCity
        End Get
        Set(ByVal value As String)
            mCity = value
        End Set
    End Property

    Public Property ZipCode() As String
        Get
            Return mZipCode
        End Get
        Set(ByVal value As String)
            mZipCode = value
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

    Public Property PaymentTerm() As Double
        Get
            Return mPaymentTerm
        End Get
        Set(ByVal value As Double)
            mPaymentTerm = value
        End Set
    End Property

    Public Property CreditLimit() As Double
        Get
            Return mCreditLimit
        End Get
        Set(ByVal value As Double)
            mCreditLimit = value
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

    Public Property NPWP_Name() As String
        Get
            Return mNPWP_Name
        End Get
        Set(ByVal value As String)
            mNPWP_Name = value
        End Set
    End Property

    Public Property DefaultTaxCode() As String
        Get
            Return mDefaultTaxCode
        End Get
        Set(ByVal value As String)
            mDefaultTaxCode = value
        End Set
    End Property

    Public Property TaxAddressLine1() As String
        Get
            Return mTaxAddressLine1
        End Get
        Set(ByVal value As String)
            mTaxAddressLine1 = value
        End Set
    End Property

    Public Property TaxAddressLine2() As String
        Get
            Return mTaxAddressLine2
        End Get
        Set(ByVal value As String)
            mTaxAddressLine2 = value
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

    Public Property PrimaryCustomerReference() As Double
        Get
            Return mPrimaryCustomerReference
        End Get
        Set(ByVal value As Double)
            mPrimaryCustomerReference = value
        End Set
    End Property

    Public Property JoinDate() As String
        Get
            Return mJoinDate
        End Get
        Set(ByVal value As String)
            mJoinDate = value
        End Set
    End Property

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

    Public Overridable Function SimpanBaru(ByVal vusercode As String, ByVal dgvCP As DataGridView, ByVal dgvAdd As DataGridView, _
    ByVal dgvBank As DataGridView, ByVal vFrm As FrmMasterCustomer) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mprimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.NoType, vusercode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vusercode, dgvCP, dgvAdd, dgvBank, vFrm)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Public Function SimpanBaru1(ByVal vusercode As String, ByVal dgvCP As DataGridView, ByVal dgvAdd As DataGridView, _
    ByVal dgvBank As DataGridView, ByVal vFrm As FrmMasterCustomer) As Boolean
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim DTA1 As New dseCustomerTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim DTA2 As New dseCustomerTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim DTA3 As New dseCustomerTableAdapters.tblmaster_customer_item_bankTableAdapter
        Dim row As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString

            If DTA.InsertQueryNew(Me.mCustomerCode, mcustomerName, mCity, mCountry, mZipCode, mCreditLimit, CInt(mPaymentTerm), mprimaryCustomerGroup, _
                mprimarySalesArea, mprimaryCustomerSegmentation, mprimaryCustomerType, mNPWP, mDefaultTaxCode, mTaxAddressLine1, mTaxAddressLine2, _
                vusercode, Now, vusercode, Now, Now, Me.mRemarks, Me.mPrimarySalesPerson, Me.mEFak_Blok, Me.mEFak_No, Me.mEFak_RT, Me.mEFak_RW, _
                Me.mEFak_Kecamatan, Me.mEFak_Kelurahan, Me.mEFak_Kabupaten, Me.mEFak_Propinsi, Me.mEFak_KodePos, Me.mEFak_Telepon, Me.mEFak_NamaJalan, _
                Me.mNoNIKPASS, Me.mOptionNIKPASS, Me.mPrimaryCustomerReference, Me.mJoinDate, Me.mNPWP_Name, Me.mNPWPLama, Me.mNPWPType, Me.mTKU_ID, Me.mEmail) > 0 Then  '//TR : Tambahan baru Nik/Pas
                Me.mprimary = DTA.ScalarQuery_GetPrimary(Me.mCustomerCode).Value
                Me.mTimestamp = DTA.ScalarQuery_getTimeStamp(CDec(Me.mprimary)).Value
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vusercode
                Me.mModifiedBy = vusercode

                'simpan data contact person
                For Each row In dgvCP.Rows
                    With row
                        If row.Index <> dgvCP.Rows.Count - 1 Then
                            If .Cells(vFrm.ColPersonSendCL.Index).Value = Boolean.TrueString Then
                                DTA1.InsertQuery(Me.mprimary, .Cells(vFrm.ColPersonName.Index).Value, .Cells(vFrm.ColPersonPosition.Index).Value, _
                                .Cells(vFrm.ColPersonPhoneNumber.Index).Value, .Cells(vFrm.ColPersonFax.Index).Value, .Cells(vFrm.ColPersonEmail.Index).Value, _
                                1, .Cells(vFrm.ColPerson_Remarks.Index).Value, CStr(Me.mprimary) + "-" + CStr(row.Index))
                            Else
                                DTA1.InsertQuery(Me.mprimary, .Cells(vFrm.ColPersonName.Index).Value, .Cells(vFrm.ColPersonPosition.Index).Value, _
                                .Cells(vFrm.ColPersonPhoneNumber.Index).Value, .Cells(vFrm.ColPersonFax.Index).Value, .Cells(vFrm.ColPersonEmail.Index).Value, _
                                0, .Cells(vFrm.ColPerson_Remarks.Index).Value, CStr(Me.mprimary) + "-" + CStr(row.Index))
                            End If
                            'DTA1.InsertQuery(Me.mprimary, CStr(.Cells(1).Value), CStr(.Cells(2).Value), CStr(.Cells(3).Value), _
                            'CStr(.Cells(4).Value), CStr(.Cells(5).Value), CStr(.Cells(6).Value), CStr(Me.mprimary) + "-" + CStr(row.Index))
                        End If
                    End With
                Next
                'simpan data address    /tambahan sales area
                For Each row In dgvAdd.Rows
                    With row
                        If row.Index <> dgvAdd.Rows.Count - 1 Then
                            DTA2.Insert(Me.mprimary, CStr(.Cells(vFrm.Column1.Index).Value), CStr(.Cells(vFrm.Col_Line1.Index).Value), _
                            CStr(.Cells(vFrm.Col_Line2.Index).Value), CStr(.Cells(vFrm.Col_Line3.Index).Value), CStr(.Cells(vFrm.Col_Line4.Index).Value), _
                            CStr(.Cells(vFrm.Col_Line5.Index).Value), CStr(.Cells(vFrm.ColPhoneAdd.Index).Value), CStr(.Cells(vFrm.ColFaxAdd.Index).Value), _
                            CStr(.Cells(vFrm.ColRemarksAdd.Index).Value), CStr(Me.mprimary) + "-" + CStr(.Index), CDbl(.Cells(vFrm.Col_PrimarySalesArea.Index).Value))
                        End If
                    End With
                Next
                'simpan data bank
                For Each row In dgvBank.Rows
                    With row
                        If row.Index <> dgvBank.Rows.Count - 1 Then
                            DTA3.Insert(Me.mprimary, CStr(.Cells(1).Value), CStr(.Cells(2).Value), CStr(.Cells(8).Value), _
                            CStr(.Cells(4).Value), CStr(.Cells(5).Value), CStr(.Cells(6).Value), CStr(.Cells(7).Value), CStr(Me.mprimary) _
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

    Public Overridable Function SimpanPerubahan(ByVal vusercode As String, ByVal dgvCP As DataGridView, ByVal dgvAdd As DataGridView, _
    ByVal dgvBank As DataGridView, ByVal vFrm As FrmMasterCustomer) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mprimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.NoType, vusercode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vusercode, dgvCP, dgvAdd, dgvBank, vFrm)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Public Function SimpanPerubahan1(ByVal vusercode As String, ByVal dgvCP As DataGridView, ByVal dgvAdd As DataGridView, _
    ByVal dgvBank As DataGridView, ByVal vFrm As FrmMasterCustomer) As Boolean
        Dim dta As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim DTA1 As New dseCustomerTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim DTA2 As New dseCustomerTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim DTA3 As New dseCustomerTableAdapters.tblmaster_customer_item_bankTableAdapter
        Dim row As DataGridViewRow

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = dta.Connection.ConnectionString
            DTA2.Connection.ConnectionString = dta.Connection.ConnectionString
            DTA3.Connection.ConnectionString = dta.Connection.ConnectionString


            If dta.UpdateQueryNew(Me.mCustomerCode, mcustomerName, mCity, mCountry, mZipCode, _
                mCreditLimit, mPaymentTerm, mprimaryCustomerGroup, mprimarySalesArea, mprimaryCustomerSegmentation, PrimaryCustomertype, _
                mNPWP, mDefaultTaxCode, mTaxAddressLine1, mTaxAddressLine2, vusercode, Now, Now, Me.mRemarks, Me.mPrimarySalesPerson, _
                Me.mEFak_Blok, Me.mEFak_No, Me.mEFak_RT, Me.mEFak_RW, Me.mEFak_Kecamatan, Me.mEFak_Kelurahan, Me.mEFak_Kabupaten, _
                Me.mEFak_Propinsi, Me.mEFak_KodePos, Me.mEFak_Telepon, Me.mEFak_NamaJalan, Me.mNoNIKPASS, Me.mOptionNIKPASS, Me.mPrimaryCustomerReference, _
                Me.mJoinDate, Me.mNPWP_Name, Me.mNPWPLama, Me.mNPWPType, Me.mTKU_ID, Me.mEmail, mprimary, mTimestamp) > 0 Then  '//TR : Tambahan baru Nik/Pas
                Me.mTimestamp = dta.ScalarQuery_getTimeStamp(Me.mprimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vusercode

                'hapus data item lama
                DTA1.DeleteQuery(Me.mprimary)
                DTA2.Delete(Me.mprimary)
                DTA3.Delete(Me.mprimary)
                'simpan data item baru
                'simpan data contact person
                For Each row In dgvCP.Rows
                    With row
                        If row.Index <> dgvCP.Rows.Count - 1 Then
                            If .Cells(vFrm.ColPersonSendCL.Index).Value = Boolean.TrueString Then
                                DTA1.InsertQuery(Me.mprimary, .Cells(vFrm.ColPersonName.Index).Value, .Cells(vFrm.ColPersonPosition.Index).Value, _
                                .Cells(vFrm.ColPersonPhoneNumber.Index).Value, .Cells(vFrm.ColPersonFax.Index).Value, .Cells(vFrm.ColPersonEmail.Index).Value, _
                                1, .Cells(vFrm.ColPerson_Remarks.Index).Value, CStr(Me.mprimary) + "-" + CStr(row.Index))
                            Else
                                DTA1.InsertQuery(Me.mprimary, .Cells(vFrm.ColPersonName.Index).Value, .Cells(vFrm.ColPersonPosition.Index).Value, _
                                .Cells(vFrm.ColPersonPhoneNumber.Index).Value, .Cells(vFrm.ColPersonFax.Index).Value, .Cells(vFrm.ColPersonEmail.Index).Value, _
                                0, .Cells(vFrm.ColPerson_Remarks.Index).Value, CStr(Me.mprimary) + "-" + CStr(row.Index))
                            End If

                            'DTA1.Insert(Me.mprimary, .Cells(1).Value, .Cells(2).Value, .Cells(3).Value, _
                            '.Cells(4).Value, .Cells(5).Value, .Cells(6).Value, CStr(Me.mprimary) + "-" + CStr(row.Index))
                        End If
                    End With
                Next
                'simpan data address    /tambahan sales area
                For Each row In dgvAdd.Rows
                    With row
                        If row.Index <> dgvAdd.Rows.Count - 1 Then
                            DTA2.Insert(Me.mprimary, CStr(.Cells(vFrm.Column1.Index).Value), CStr(.Cells(vFrm.Col_Line1.Index).Value), _
                            CStr(.Cells(vFrm.Col_Line2.Index).Value), CStr(.Cells(vFrm.Col_Line3.Index).Value), CStr(.Cells(vFrm.Col_Line4.Index).Value), _
                            CStr(.Cells(vFrm.Col_Line5.Index).Value), CStr(.Cells(vFrm.ColPhoneAdd.Index).Value), CStr(.Cells(vFrm.ColFaxAdd.Index).Value), _
                            CStr(.Cells(vFrm.ColRemarksAdd.Index).Value), CStr(Me.mprimary) + "-" + CStr(.Index), CDbl(.Cells(vFrm.Col_PrimarySalesArea.Index).Value))
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

    Public Overridable Function Hapus(ByVal vprimary As Double, Optional ByVal vUserCode As String = "") As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mprimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.NoType, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vprimary)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Public Function Hapus1(ByVal vprimary As Double) As Boolean
        Dim dta As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

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

    Public Function ApakahSudahAda_CP(ByVal vPrimaryItemCP As Double, ByVal vNamaCP As String, ByVal vPrimaryCust As Double) As Boolean
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer_itemcpDataTable
        Dim row As dseCustomer.tblmaster_customer_itemcpRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryCustomer(TBL, vPrimaryCust)
            For Each row In TBL.Rows
                '1. cek apakah primary item-nya ada
                If row.Primary = vPrimaryItemCP Then
                    Return True
                End If
                '2. cek apakah kombinasi nama dan primary custnya ada
                If row.Name.Trim = vNamaCP Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vCustomerCode As String) As Boolean
        Dim Dta As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Dta.ScalarQuery_GetjmlCustCode(vCustomerCode).HasValue Then
                If Dta.ScalarQuery_GetjmlCustCode(vCustomerCode).Value > 0 Then
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

    Public Function ApakahBolehHapus_Address(ByVal v_Identifier As String) As Boolean
        Dim DTA_Quo As New dseCustomerTableAdapters.tbltrans_quotationTableAdapter
        Try
            'cek apakah sudah digunakan pada SO, DO, Invoice, Quotation 
            'jika sudah digunakan pada Quotation, code no setting, SO, Cust Complain
            DTA_Quo.Connection.ConnectionString = Core_Module.GetConstring
            If DTA_Quo.ScalarQuery_GetJmlByAlamatCust(v_Identifier).HasValue = True Then
                If DTA_Quo.ScalarQuery_GetJmlByAlamatCust(v_Identifier).Value > 0 Then
                    MsgBox("This data is has been used, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Dim DTA_SO As New dseCustomerTableAdapters.tbltrans_soTableAdapter
            DTA_SO.Connection.ConnectionString = Core_Module.GetConstring
            If DTA_SO.ScalarQuery_GetJmlByIdentifierAlamat(v_Identifier).HasValue = True Then
                If DTA_SO.ScalarQuery_GetJmlByIdentifierAlamat(v_Identifier).Value > 0 Then
                    MsgBox("This data is has been used, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            DTA_SO.Dispose()

            Dim DTA_Inv As New dseCustomerTableAdapters.tbltrans_salesinvoiceTableAdapter
            DTA_Inv.Connection.ConnectionString = Core_Module.GetConstring
            If DTA_Inv.ScalarQuery_GetJmlByCustAddress(v_Identifier).HasValue = True Then
                If DTA_Inv.ScalarQuery_GetJmlByCustAddress(v_Identifier).Value > 0 Then
                    MsgBox("This data is has been used, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            DTA_Inv.Dispose()

            Dim DTA_Inv_OB As New dseCustomerTableAdapters.tbltrans_salesinvoice_obseleteTableAdapter
            DTA_Inv_OB.Connection.ConnectionString = Core_Module.GetConstring
            If DTA_Inv_OB.ScalarQuery_GEtJml(v_Identifier).HasValue = True Then
                If DTA_Inv_OB.ScalarQuery_GEtJml(v_Identifier).Value > 0 Then
                    MsgBox("This data is has been used, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            DTA_Inv_OB.Dispose()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Quo.Dispose()

        End Try
    End Function

    Public Function ApakahFormatNPWPBenar(ByVal Value As String, Optional ByVal ShowMsg As Boolean = False) As Boolean
        Dim Reference As String = "00.000.000.0-000.000"
        Dim Dummy As String = ""

        Try
            If Value.Length <> Reference.Length Then
                If ShowMsg = True Then
                    MsgBox("Please specify a valid NPWP format", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'jika mengandung spasi
            If Value.Contains(" ") = True Then
                If ShowMsg = True Then
                    MsgBox("Please specify a valid NPWP format", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'semua harus angka, kecuali titik dan -
            Dummy = Value.Substring(0, 2)
            If IsNumeric(Dummy) = False Then
                If ShowMsg = True Then
                    MsgBox("Please specify a valid NPWP format", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Dummy = Value.Substring(2, 1)
            If Dummy <> "." Then
                If ShowMsg = True Then
                    MsgBox("Please specify a valid NPWP format", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Dummy = Value.Substring(3, 3)
            If IsNumeric(Dummy) = False Then
                If ShowMsg = True Then
                    MsgBox("Please specify a valid NPWP format", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Dummy = Value.Substring(6, 1)
            If Dummy <> "." Then
                If ShowMsg = True Then
                    MsgBox("Please specify a valid NPWP format", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Dummy = Value.Substring(7, 3)
            If IsNumeric(Dummy) = False Then
                If ShowMsg = True Then
                    MsgBox("Please specify a valid NPWP format", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Dummy = Value.Substring(10, 1)
            If Dummy <> "." Then
                If ShowMsg = True Then
                    MsgBox("Please specify a valid NPWP format", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Dummy = Value.Substring(11, 1)
            If IsNumeric(Dummy) = False Then
                If ShowMsg = True Then
                    MsgBox("Please specify a valid NPWP format", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Dummy = Value.Substring(12, 1)
            If Dummy <> "-" Then
                If ShowMsg = True Then
                    MsgBox("Please specify a valid NPWP format", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Dummy = Value.Substring(13, 3)
            If IsNumeric(Dummy) = False Then
                If ShowMsg = True Then
                    MsgBox("Please specify a valid NPWP format", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Dummy = Value.Substring(16, 1)
            If Dummy <> "." Then
                If ShowMsg = True Then
                    MsgBox("Please specify a valid NPWP format", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Dummy = Value.Substring(17, 3)
            If IsNumeric(Dummy) = False Then
                If ShowMsg = True Then
                    MsgBox("Please specify a valid NPWP format", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vprimary As Double) As Boolean
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim DTA1 As New dseMasterSupplierTableAdapters.tblmaster_supplierTableAdapter

        Dim TipeDPInv As Integer = 0
        Dim Jml As Nullable(Of Long)
        Dim TipeTransaksiCN As Integer = 0
        Dim TipeTransaksiDN As Integer = 1
        Dim TipeKartuCNDN As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            'msgbox("belum")
            'jika sudah digunakan pada Quotation, code no setting, SO, Cust Complain
            Dim DTA_Quo As New dseCustomerTableAdapters.tbltrans_quotationTableAdapter
            DTA_Quo.Connection.ConnectionString = Core_Module.GetConstring
            If DTA_Quo.ScalarQuery_GetJmlByPrimaryCust(vprimary).HasValue = True Then
                If DTA_Quo.ScalarQuery_GetJmlByPrimaryCust(vprimary).Value > 0 Then
                    MsgBox("This data is refered in quotation data, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            DTA_Quo.Dispose()

            Dim DTA_SO As New dseCustomerTableAdapters.tbltrans_soTableAdapter
            DTA_SO.Connection.ConnectionString = Core_Module.GetConstring
            If DTA_SO.ScalarQuery_GetJml(vprimary).HasValue = True Then
                If DTA_SO.ScalarQuery_GetJml(vprimary).Value > 0 Then
                    MsgBox("This data is refered in sales order data, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            DTA_SO.Dispose()

            'Dim DTA_CodeSetting As New dseCustomerTableAdapters.tbltrans_soTableAdapter
            'DTA_CodeSetting.Connection.ConnectionString = Core_Module.GetConstring
            'If DTA_CodeSetting.ScalarQuery_GetJml(vprimary).HasValue = True Then
            '    If DTA_CodeSetting.ScalarQuery_GetJml(vprimary).Value > 0 Then
            '        MsgBox("This data is refered in code no setting data, you cannot delete this data.", _
            '        MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If
            'DTA_CodeSetting.Dispose()

            'Dim DTA_CustComplain As New dseCustomerTableAdapters.tbltrans_soTableAdapter
            'DTA_CustComplain.Connection.ConnectionString = Core_Module.GetConstring
            'If DTA_CustComplain.ScalarQuery_GetJml(vprimary).HasValue = True Then
            '    If DTA_CustComplain.ScalarQuery_GetJml(vprimary).Value > 0 Then
            '        MsgBox("This data is refered in customer complain data, you cannot delete this data.", _
            '        MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If
            'DTA_CustComplain.Dispose()

            'inv dp
            Jml = DTA.ScalarQuery_GetJmlInv(vprimary, TipeDPInv)
            If Jml.HasValue = True Then
                If Jml.Value > 0 Then
                    MsgBox("This data is refered downpayment invoice data, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'credit - note
            Jml = DTA1.ScalarQuery_GetJmlCNDN(vprimary, TipeTransaksiCN, TipeKartuCNDN)
            If Jml.HasValue = True Then
                If Jml.Value > 0 Then 
                    MsgBox("There are credit notes refering to this customer. You cannot delete this data", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If 
            End If

            'debit - note 
            Jml = DTA1.ScalarQuery_GetJmlCNDN(vprimary, TipeTransaksiDN, TipeKartuCNDN)
            If Jml.HasValue = True Then
                If Jml.Value > 0 Then
                    MsgBox("There are debit notes refering to this customer. You cannot delete this data", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'price list
            Jml = DTA.ScalarQuery_GetJmlPriceList(vprimary)
            If Jml.HasValue = True Then
                If Jml.Value > 0 Then
                    MsgBox("This data is refered by product price list data, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'quick sales return 
            Jml = DTA.ScalarQuery_GetJmlSRQuick(1, vprimary)
            If Jml.HasValue = True Then
                If Jml.Value > 0 Then
                    MsgBox("This data is refered by quick sales return data, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'begining balance sales
            Dim DTA_SO1 As New dseCustomerTableAdapters.tbltrans_soTableAdapter
            DTA_SO1.Connection.ConnectionString = Core_Module.GetConstring
            If DTA_SO1.ScalarQuery_GetJmlBBSales(vprimary).HasValue = True Then
                If DTA_SO1.ScalarQuery_GetJmlBBSales(vprimary).Value > 0 Then
                    MsgBox("This data is refered in sales begining balance, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'begining balance sales return
            If DTA_SO1.ScalarQuery_GetJmlBBSalesReturn(vprimary).HasValue = True Then
                If DTA_SO1.ScalarQuery_GetJmlBBSalesReturn(vprimary).Value > 0 Then
                    MsgBox("This data is refered in sales return begining balance, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            DTA_SO1.Dispose()

            ''begining balance DP -> ga perlu, karena untuk bikin begining balance dp, pasti ada begining balance so
            'If DTA_SO1.ScalarQuery_GetJmlBB_DP(vprimary, 1).HasValue = True Then
            '    If DTA_SO1.ScalarQuery_GetJmlBB_DP(vprimary, 1).Value > 0 Then
            '        MsgBox("This data is refered in sales DP begining balance, you cannot delete this data.", _
            '        MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If


            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlBy_CustSegmentation(ByVal vPrimaryCustSegmentation As Double) As Integer
        Dim dta As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.ScalarQuery_GetJmlPrimaryCustByPrimaryCustSegmen(vPrimaryCustSegmentation).HasValue Then
                Return dta.ScalarQuery_GetJmlPrimaryCustByPrimaryCustSegmen(vPrimaryCustSegmentation).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetJmlBy_CustGroup(ByVal vPrimaryCustGroup As Double) As Integer
        Dim dta As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.ScalarQuery_GetjmlPrimaryCustByPrimaryCustGroup(vPrimaryCustGroup).HasValue Then
                Return dta.ScalarQuery_GetjmlPrimaryCustByPrimaryCustGroup(vPrimaryCustGroup).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetJmlBy_CustType(ByVal vPrimaryCustType As Double) As Integer
        Dim dta As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.ScalarQuery_GetJmlPrimaryCustByPrimaryCustType(vPrimaryCustType).HasValue Then
                Return dta.ScalarQuery_GetJmlPrimaryCustByPrimaryCustType(vPrimaryCustType).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetJmlBy_CustSalesArea(ByVal vPrimarySalesArea As Double) As Integer
        Dim dta As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.ScalarQuery_GetJmlPrimaryCustByPrimarySales(vPrimarySalesArea).HasValue Then
                Return dta.ScalarQuery_GetJmlPrimaryCustByPrimarySales(vPrimarySalesArea).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetPrimary_SalesPerson(ByVal vPrimaryCust As Double) As Double
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimarySalesPerson(vPrimaryCust)
            If RetVal.HasValue Then
                Return RetVal.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary_CustGroup(ByVal vPrimaryCust As Double) As Double
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimaryCustGroup(vPrimaryCust)
            If RetVal.HasValue Then
                Return RetVal.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary_CustSegmentasi(ByVal vPrimaryCust As Double) As Double
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimarySegmentation(vPrimaryCust)
            If RetVal.HasValue Then
                Return RetVal.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary_CustType(ByVal vPrimaryCust As Double) As Double
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimaryCustTipe(vPrimaryCust)
            If RetVal.HasValue Then
                Return RetVal.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary_CustSalesArea(ByVal vPrimaryCust As Double) As Double
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimarySalesArea(vPrimaryCust)
            If RetVal.HasValue Then
                Return RetVal.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal CustCode As String) As Double
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(CustCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(CustCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPhoneNo_CP(ByVal vPrimaryCust As Double, ByVal vCPIdentifier As String) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer_itemcpDataTable
        Dim row As dseCustomer.tblmaster_customer_itemcpRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_CustomerNIdentifier(TBL, vPrimaryCust, vCPIdentifier) > 0 Then
                row = TBL.Rows(0)
                Return row.Phone_No
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

    Public Function GetPhoneNo(ByVal vPrimaryCust As Double, ByVal vAddIdentifier As String) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer_item_addressDataTable
        Dim row As dseCustomer.tblmaster_customer_item_addressRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_CustomerNIdentifier(TBL, vPrimaryCust, vAddIdentifier) > 0 Then
                row = TBL.Rows(0)
                Return row.Phone_No
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

    Public Function GetPaymentTerm(ByVal vPrimaryCust As Double) As Double
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPaymentTerms(vPrimaryCust).HasValue = True Then
                Return DTA.ScalarQuery_GetPaymentTerms(vPrimaryCust).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub GetData_EFaktur(ByVal vPrimaryCust As Double, ByVal vTxtNamaJalan As TextBox, ByVal vTxtBlok As TextBox, ByVal vTxtNo As TextBox, _
    ByVal vTxtRT As TextBox, ByVal vTxtRW As TextBox, ByVal vTxtKelurahan As TextBox, ByVal vTxtKecamatan As TextBox, ByVal vTxtKabupaten As TextBox, _
    ByVal vTxtPropinsi As TextBox, ByVal vTxtKodePos As TextBox, ByVal vTxtNoTelepon As TextBox)
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customerDataTable
        Dim row As dseCustomer.tblmaster_customerRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, vPrimaryCust)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                vTxtBlok.Text = row.Blok
                vTxtKabupaten.Text = row.Kabupaten
                vTxtKecamatan.Text = row.Kecamatan
                vTxtKelurahan.Text = row.Kelurahan
                vTxtKodePos.Text = row.Kode_Pos
                vTxtNamaJalan.Text = row.Nama_Jalan
                vTxtNo.Text = row.Nomor
                vTxtNoTelepon.Text = row.No_Telepon
                vTxtPropinsi.Text = row.Propinsi
                vTxtRT.Text = row.RT
                vTxtRW.Text = row.RW
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetNPWP(ByVal vPrimaryCust As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNPWP(vPrimaryCust)) = False Then
                Return DTA.ScalarQuery_GetNPWP(vPrimaryCust)
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

    Public Function GetTaxAddLine1(ByVal vPrimaryCust As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetTaxAddLine1(vPrimaryCust)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetTaxAddLine1(vPrimaryCust)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTaxAddLine2(ByVal vPrimaryCust As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetTaxAddLine2(vPrimaryCust)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetTaxAddLine2(vPrimaryCust)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCustName(ByVal vPrimary As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCustName(vPrimary)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetCustName(vPrimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetAddress(ByVal vIdentifierAdd As String, ByVal vPrimaryCustomer As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer_item_addressDataTable
        Dim row As dseCustomer.tblmaster_customer_item_addressRow
        Dim RetVal As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_CustomerNIdentifier(TBL, vPrimaryCustomer, vIdentifierAdd) > 0 Then
                row = TBL.Rows(0)

                RetVal = row.Line_1
                RetVal = RetVal + vbCrLf + row.Line_2
                RetVal = RetVal + vbCrLf + row.Line_3
                RetVal = RetVal + vbCrLf + row.Line_4
                RetVal = RetVal + vbCrLf + row.Line_5
            End If

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function Get_City_ByPrimary(ByVal vPrimaryCust As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCity(vPrimaryCust)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetCity(vPrimaryCust)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub GetAlamatPajak(ByVal vPrimaryCust As Double, ByRef vAlamat1 As String, ByRef vAlamat2 As String)
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customerDataTable
        Dim row As dseCustomer.tblmaster_customerRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_Primary(TBL, vPrimaryCust) > 0 Then
                row = TBL.Rows(0)
                vAlamat1 = row.Tax_Add_Line1
                vAlamat2 = row.Tax_Add_Line2
            Else
                vAlamat1 = ""
                vAlamat2 = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetCustContactPerson(ByVal vPrimaryCust As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer_itemcpDataTable
        Dim row As dseCustomer.tblmaster_customer_itemcpRow
        Dim ans As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_PrimaryCustomer(TBL, vPrimaryCust) > 0 Then
                For Each row In TBL.Rows
                    If ans = "" Then
                        ans = row.Name
                    Else
                        ans = ans + ", " + row.Name
                    End If
                Next
            Else
                Return ""
            End If

            Return ans
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetCustContactPerson_ForCollectionLetter(ByVal vPrimaryCust As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer_itemcpDataTable
        Dim row As dseCustomer.tblmaster_customer_itemcpRow
        Dim ans As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_PrimaryCustomer(TBL, vPrimaryCust) > 0 Then
                For Each row In TBL.Rows
                    If row.Send_Collection_Letter = 1 Then
                        If ans = "" Then
                            ans = row.Name
                        Else
                            ans = ans + "/" + row.Name
                        End If
                    End If
                Next
            Else
                Return ""
            End If

            Return ans
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetCreditLimit(ByVal vPrimaryCust As Double) As Decimal
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim Dummy As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetCreditLimit(vPrimaryCust)
            If Dummy.HasValue = True Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    '//TR : Tambahan baru Nik/Pas
    Public Function GetPilihNikPass(ByVal vPrimaryCust As Double) As Integer
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim Dummy As Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPilihNikPassByPrimary(vPrimaryCust)
            If Dummy.HasValue = True Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetNoNikPass(ByVal vPrimaryCust As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoNikPassByPrimary(vPrimaryCust)) = False Then
                Return DTA.ScalarQuery_GetNoNikPassByPrimary(vPrimaryCust)
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
    Public Function GetNoNikPass2(ByVal vPrimaryCust As Double, ByVal vPilih As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoNikPassByPrimaryandPilih(vPrimaryCust, vPilih)) = False Then
                Return DTA.ScalarQuery_GetNoNikPassByPrimaryandPilih(vPrimaryCust, vPilih)
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
    Public Function GetCustomerCode(ByVal vPrimary As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCustCode(vPrimary)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetCustCode(vPrimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlBy_PrimaryCurrency(ByVal vPrimaryCurrency As Double) As Double
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_item_bankTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vPrimaryCurrency).HasValue = True Then
                Return DTA.ScalarQuery_GetJml(vPrimaryCurrency).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKodeCust As String) As Double
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetjmlCustCode(vKodeCust).HasValue = True Then
                Return DTA.ScalarQuery_GetjmlCustCode(vKodeCust).Value
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
        Dim DTA As New dseCustomerTableAdapters.DataTable1TableAdapter

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
            ElseIf vSearchBy = EnumSearchCriteria.BySalesPCode Then
                If DTA.ScalarQuery_GetJml4(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml4(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.BySalesPName Then
                If DTA.ScalarQuery_GetJml5(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml5(vKriteria + "%").Value
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

    Public Function IsiListBoxCustomer(ByVal lbo As CheckedListBox, ByVal TBL As dseCustomer.tblmaster_customer1DataTable) As Boolean
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer1TableAdapter
        Dim row As dseCustomer.tblmaster_customer1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            lbo.DataSource = TBL.DefaultView
            For Each row In TBL.Rows
                row.Combination = row.Customer_Code + " - " + row.Customer_Name
            Next
            lbo.DisplayMember = "Combination"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetKodeAddress(ByVal vIdentifierAdd As String) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_item_addressTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetAddCode(vIdentifierAdd)) = False Then
                Return DTA.ScalarQuery_GetAddCode(vIdentifierAdd)
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

    Public Function GetTaxCode(ByVal vPrimaryCust As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDefaultTaxCode(vPrimaryCust)) = False Then
                Return DTA.ScalarQuery_GetDefaultTaxCode(vPrimaryCust)
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

    Public Function GetIdentifierAdd(ByVal vPrimaryCust As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer_item_addressDataTable
        Dim row As dseCustomer.tblmaster_customer_item_addressRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryCustomer(TBL, vPrimaryCust)
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

    Public Function GetIdentifier_CP(ByVal vPrimaryCust As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer_itemcpDataTable
        Dim row As dseCustomer.tblmaster_customer_itemcpRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryCustomer(TBL, vPrimaryCust)
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

    Public Function GetIdentifierAdd(ByVal vPrimaryCust As Double, ByVal vKodeAdd As String) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_item_addressTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetIdentifier(vPrimaryCust, vKodeAdd)) = False Then
                Return DTA.ScalarQuery_GetIdentifier(vPrimaryCust, vKodeAdd)
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

    Public Function IsiTxtAddress(ByVal vIdentifierAdd As String, ByVal vPrimaryCustomer As Double, ByVal txt As TextBox) As Boolean
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer_item_addressDataTable
        Dim row As dseCustomer.tblmaster_customer_item_addressRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_CustomerNIdentifier(TBL, vPrimaryCustomer, vIdentifierAdd) > 0 Then
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

    Public Function IsiCboAddress(ByVal cbo As ComboBox, ByVal PrimaryCustomer As Double) As Boolean
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer_item_addressDataTable
        Dim row As dseCustomer.tblmaster_customer_item_addressRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            cbo.Items.Clear()
            If DTA.FillBy_PrimaryCustomer(TBL, PrimaryCustomer) > 0 Then
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

    Public Function IsiCboAddress(ByVal vCbo As ComboBox, ByVal vPrimaryCustomer As Double, ByVal TBL As dseCustomer.tblmaster_customer_item_addressDataTable) As Boolean
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_item_addressTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vCbo.DataSource = Nothing
            vCbo.Items.Clear()

            DTA.FillBy_PrimaryCustomer(TBL, vPrimaryCustomer)
            vCbo.DataSource = TBL.DefaultView
            vCbo.DisplayMember = "Address Code"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
     ByVal DGV As DataGridView, ByVal TBL As dseCustomer.DataTable1DataTable, _
     ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New dseCustomerTableAdapters.DataTable1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByAccountNo Then
                QueryResult = DTA.FillBy_AccNo(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPCode Then
                QueryResult = DTA.FillBy_SalesPersonCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPName Then
                QueryResult = DTA.FillBy_SalesPersonNameLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Customer code"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
            End If


            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_getTimeStamp(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal dgvAddress As DataGridView, _
    ByVal dgvItemCP As DataGridView, ByVal dgvItemBank As DataGridView) As Boolean
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer3TableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer3DataTable
        Dim row As dseCustomer.tblmaster_customer3Row
        Dim DTA1 As New dseCustomerTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim TBL1 As New dseCustomer.tblmaster_customer_item_addressDataTable
        Dim row1 As dseCustomer.tblmaster_customer_item_addressRow
        Dim DTA2 As New dseCustomerTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim TBL2 As New dseCustomer.tblmaster_customer_itemcpDataTable
        Dim row2 As dseCustomer.tblmaster_customer_itemcpRow
        Dim DTA3 As New dseCustomerTableAdapters.tblmaster_customer_item_bankTableAdapter
        Dim TBL3 As New dseCustomer.tblmaster_customer_item_bankDataTable
        Dim row3 As dseCustomer.tblmaster_customer_item_bankRow
        Dim Dummy1 As Boolean = False

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.Fill(TBL, vPrimary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mCity = .City
                    Me.mCountry = .Country
                    Me.mCreditLimit = .Credit_Limit
                    Me.mCustomerCode = .Customer_Code
                    Me.mcustomerName = .Customer_Name
                    Me.mDefaultTaxCode = .Default_Tax_Code
                    Me.mNPWP = .NPWP_No
                    Me.mNPWP_Name = .NPWP_Name
                    Me.mPaymentTerm = .Payment_Terms
                    Me.mprimary = .Primary
                    Me.mprimaryCustomerGroup = .Primary_Customer_Group
                    Me.mprimaryCustomerSegmentation = .Primary_Customer_Segmentasi
                    Me.mprimaryCustomerType = .Primary_Customer_Type
                    Me.mprimarySalesArea = .Primary_Sales_Area
                    Me.mTaxAddressLine1 = .Tax_Add_Line1
                    Me.mTaxAddressLine2 = .Tax_Add_Line2
                    Me.mZipCode = .ZIP_Code
                    Me.mRemarks = .Remarks
                    Me.mPrimarySalesPerson = .primary_sales_person

                    Me.mEFak_Blok = .Blok
                    Me.mEFak_Kabupaten = .Kabupaten
                    Me.mEFak_Kecamatan = .Kecamatan
                    Me.mEFak_Kelurahan = .Kelurahan
                    Me.mEFak_KodePos = .Kode_Pos
                    Me.mEFak_NamaJalan = .Nama_Jalan
                    Me.mEFak_No = .Nomor
                    Me.mEFak_Propinsi = .Propinsi
                    Me.mEFak_RT = .RT
                    Me.mEFak_RW = .RW
                    Me.mEFak_Telepon = .No_Telepon
                    Me.mPrimaryCustomerReference = .Primary_Customer_Reference
                    Me.mJoinDate = .Join_Date

                    '//TR : Tambahan baru Nik/Pas
                    Me.mOptionNIKPASS = .Pilih_NIKPASS
                    Me.mNoNIKPASS = .No_NIKPASS

                    'update NPWP 16 Digit
                    Me.mNPWPLama = .NPWP_Old
                    Me.mNPWPType = .NPWP_Type

                    'update PPN 12%
                    Me.mTKU_ID = .TKU_ID
                    Me.mEmail = .Email

                    'salin data addres
                    dgvAddress.Rows.Clear()
                    dgvItemBank.Rows.Clear()
                    dgvItemCP.Rows.Clear()
                    DTA1.FillBy_PrimaryCustomer(TBL1, Me.mprimary)
                    'aul perubahan
                    Dim Obj2 As New ClsCekData
                    Dim ObjCustSalesArea As New Core_Module.clsMasterCustomerSalesArea
                    For Each row1 In TBL1.Rows
                        With row1
                            dgvAddress.Rows.Add("-", .Address_Code, ObjCustSalesArea.GetSalesAreaCode(.primary_sales_area), _
                            ObjCustSalesArea.GetSalesAreaDesc(.primary_sales_area), .Line_1, .Line_2, .Line_3, .Line_4, .Line_5, _
                            .Phone_No, .Fax_No, .Remark, .Identifier, .primary_sales_area)
                        End With
                    Next
                    '----

                    'salin data contact person
                    DTA2.FillBy_PrimaryCustomer(TBL2, Me.mprimary)
                    For Each row2 In TBL2.Rows
                        With row2
                            If .Send_Collection_Letter = 1 Then
                                Dummy1 = True
                                dgvItemCP.Rows.Add("-", .Name, .Position, .Phone_No, .Fax_No, _
                                ._E_mail_Address, Dummy1, .Remark, .Identifier)
                            Else
                                Dummy1 = False
                                dgvItemCP.Rows.Add("-", .Name, .Position, .Phone_No, .Fax_No, _
                                ._E_mail_Address, Dummy1, .Remark, .Identifier)
                            End If
                        End With
                    Next
                    'salin data bank
                    Dim ObjCurrency As New clsMasterCurrency
                    DTA3.FillBy_PrimaryCustomer(TBL3, Me.mprimary)
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

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .TimeStamp
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

    Public Function GetEmailCp(ByVal vPrimaryCust As Double, ByVal vAddIdentifier As String) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer_itemcpDataTable
        Dim row As dseCustomer.tblmaster_customer_itemcpRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_CustomerNIdentifier(TBL, vPrimaryCust, vAddIdentifier) > 0 Then
                row = TBL.Rows(0)
                Return row._E_mail_Address
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

    Public Function IsiStringAddress(ByVal vIdentifierAdd As String, ByVal vPrimaryCustomer As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer_item_addressDataTable
        Dim row As dseCustomer.tblmaster_customer_item_addressRow
        Dim TxtString As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_CustomerNIdentifier(TBL, vPrimaryCustomer, vIdentifierAdd) > 0 Then
                row = TBL.Rows(0)
                TxtString = ""
                TxtString = row.Line_1
                TxtString = TxtString + vbCrLf + row.Line_2
                TxtString = TxtString + vbCrLf + row.Line_3
                TxtString = TxtString + vbCrLf + row.Line_4
                TxtString = TxtString + vbCrLf + row.Line_5
            End If
            Return TxtString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Sub GetEmailCpList(ByVal vPrimaryCust As Double, ByRef vListEmail As List(Of String))
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer_itemcpDataTable
        Dim row As dseCustomer.tblmaster_customer_itemcpRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vListEmail = New List(Of String)
            If DTA.FillBy_PrimaryCustomerAndSendCLYes(TBL, vPrimaryCust) > 0 Then
                Dim Expression As New System.Text.RegularExpressions.Regex("\S+@\S+\.\S+")
                For Each row In TBL
                    'hanya menambahkan data email yang valid
                    If row._E_mail_Address <> "" Then
                        If Expression.IsMatch(row._E_mail_Address) = True Then
                            vListEmail.Add(row._E_mail_Address)
                        End If
                    End If
                Next
                Expression = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Sub SetConnectionString(ByVal DTA As dseCustomerTableAdapters.tblmaster_customer_itemcp1TableAdapter)
        DTA.Connection.ConnectionString = Core_Module.GetConstring
    End Sub

    Public Sub Update_EFaktur(ByVal vPrimaryCust As Double, ByVal vUserName As String, ByVal vNamaJalan As String, ByVal vBlok As String, ByVal vNomor As String, _
    ByVal vRT As String, ByVal vRW As String, ByVal vKelurahan As String, ByVal vKecamatan As String, ByVal vKabupaten As String, ByVal vPropinsi As String, _
    ByVal vKodePos As String, ByVal vNoTelepon As String)
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customerDataTable
        Dim row As dseCustomer.tblmaster_customerRow
        Dim ObjUser As New ClsUser

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cek apakah memiliki hak akses
            If ObjUser.ApakahMemilikiHakAkses(vUserName, "FrmMasterCustomer") = True Then
                'cek apakah data e faktur di customer kosong
                DTA.FillBy_Primary(TBL, vPrimaryCust)

                'update ke dalam database
                If TBL.Rows.Count > 0 Then
                    row = TBL.Rows(0)
                    If row.Nama_Jalan = "" And row.Blok = "" And row.Nomor = "" And row.RT = "" _
                    And row.RW = "" And row.Kelurahan = "" And row.Kecamatan = "" And row.Kabupaten = "" _
                    And row.Propinsi = "" And row.Kode_Pos = "" And row.No_Telepon = "" Then
                        If MsgBox("Update e-Faktur data to customer database?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            DTA.UpdateQuery_SetData_eFaktur(vBlok, vNomor, vRT, vRW, vKecamatan, vKelurahan, vKabupaten, _
                            vPropinsi, vKodePos, vNoTelepon, vNamaJalan, vPrimaryCust)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjUser = Nothing
        End Try
    End Sub

    'Update NPWP 16 Digit
    Public Function GetNPWPOld(ByVal vPrimaryCust As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer3TableAdapter
        Dim retval As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            retval = DTA.ScalarQuery_GetNPWPOld(vPrimaryCust)
            If String.IsNullOrEmpty(retval) = False Then
                Return retval
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
