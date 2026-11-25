Option Explicit On
Option Strict Off

Public Class clsMasterCustomer
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
    Private mDefaultTaxCode As String = ""
    Private mTaxAddressLine1 As String = ""
    Private mTaxAddressLine2 As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

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

    Public Function SimpanBaru(ByVal vusercode As String, ByVal dgvCP As DataGridView, ByVal dgvAdd As DataGridView, _
    ByVal dgvBank As DataGridView) As Boolean
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim DTA1 As New dseCustomerTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim DTA2 As New dseCustomerTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim DTA3 As New dseCustomerTableAdapters.tblmaster_customer_item_bankTableAdapter
        Dim row As DataGridViewRow

        Try
            If DTA.Insert(Me.mCustomerCode, mcustomerName, mCity, mCountry, mZipCode, mCreditLimit, CInt(mPaymentTerm), mprimaryCustomerGroup, _
                mprimarySalesArea, mprimaryCustomerSegmentation, mprimaryCustomerType, mNPWP, mDefaultTaxCode, mTaxAddressLine1, mTaxAddressLine2, _
                vusercode, Now, vusercode, Now, Now) > 0 Then
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
                            DTA1.Insert(Me.mprimary, CStr(.Cells(1).Value), CStr(.Cells(2).Value), CStr(.Cells(3).Value), _
                            CStr(.Cells(4).Value), CStr(.Cells(5).Value), CStr(.Cells(6).Value), CStr(Me.mprimary) + "-" + CStr(row.Index))
                        End If
                    End With
                Next
                'simpan data address
                For Each row In dgvAdd.Rows
                    With row
                        If row.Index <> dgvAdd.Rows.Count - 1 Then
                            DTA2.Insert(Me.mprimary, CStr(.Cells(1).Value), CStr(.Cells(2).Value), CStr(.Cells(3).Value), _
                             CStr(.Cells(4).Value), CStr(.Cells(5).Value), CStr(.Cells(6).Value), CStr(.Cells(7).Value), CStr(.Cells(8).Value), _
                             CStr(.Cells(9).Value), CStr(Me.mprimary) + "-" + CStr(.Index))
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
        Dim dta As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim DTA1 As New dseCustomerTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim DTA2 As New dseCustomerTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim DTA3 As New dseCustomerTableAdapters.tblmaster_customer_item_bankTableAdapter
        Dim row As DataGridViewRow

        Try
            If dta.Update(Me.mCustomerCode, mcustomerName, mCity, mCountry, mZipCode, _
                mCreditLimit, mPaymentTerm, mprimaryCustomerGroup, mprimarySalesArea, mprimaryCustomerSegmentation, PrimaryCustomertype, _
                mNPWP, mDefaultTaxCode, mTaxAddressLine1, mTaxAddressLine2, vusercode, Now, Now, mprimary, mTimestamp) > 0 Then
                Me.mTimestamp = dta.ScalarQuery_getTimeStamp(Me.mprimary)
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
        End Try
    End Function

    Public Function Hapus(ByVal vprimary As Double) As Boolean
        Dim dta As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Try
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

    Public Function ApakahSudahAda(ByVal vCustomerCode As String) As Boolean
        Dim Dta As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Try
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
        Try
            'cek apakah sudah digunakan pada SO, DO, Invoice, Quotation 
            'jika sudah digunakan pada Quotation, code no setting, SO, Cust Complain
            Dim DTA_Quo As New dseCustomerTableAdapters.tbltrans_quotationTableAdapter
            If DTA_Quo.ScalarQuery_GetJmlByAlamatCust(v_Identifier).HasValue = True Then
                If DTA_Quo.ScalarQuery_GetJmlByAlamatCust(v_Identifier).Value > 0 Then
                    MsgBox("This data is has been used, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Dim DTA_SO As New dseCustomerTableAdapters.tbltrans_soTableAdapter
            If DTA_SO.ScalarQuery_GetJmlByIdentifierAlamat(v_Identifier).HasValue = True Then
                If DTA_SO.ScalarQuery_GetJmlByIdentifierAlamat(v_Identifier).Value > 0 Then
                    MsgBox("This data is has been used, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            DTA_SO.Dispose()

            Dim DTA_Inv As New dseCustomerTableAdapters.tbltrans_salesinvoiceTableAdapter
            If DTA_Inv.ScalarQuery_GetJmlByCustAddress(v_Identifier).HasValue = True Then
                If DTA_Inv.ScalarQuery_GetJmlByCustAddress(v_Identifier).Value > 0 Then
                    MsgBox("This data is has been used, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Dim DTA_Inv_OB As New dseCustomerTableAdapters.tbltrans_salesinvoice_obseleteTableAdapter
            If DTA_Inv_OB.ScalarQuery_GEtJml(v_Identifier).HasValue = True Then
                If DTA_Inv_OB.ScalarQuery_GEtJml(v_Identifier).Value > 0 Then
                    MsgBox("This data is has been used, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            DTA_Quo.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vprimary As Double) As Boolean
        Try
            'msgbox("belum")
            'jika sudah digunakan pada Quotation, code no setting, SO, Cust Complain
            Dim DTA_Quo As New dseCustomerTableAdapters.tbltrans_quotationTableAdapter
            If DTA_Quo.ScalarQuery_GetJmlByPrimaryCust(vprimary).HasValue = True Then
                If DTA_Quo.ScalarQuery_GetJmlByPrimaryCust(vprimary).Value > 0 Then
                    MsgBox("This data is has been used, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            DTA_Quo.Dispose()

            Dim DTA_SO As New dseCustomerTableAdapters.tbltrans_soTableAdapter
            If DTA_SO.ScalarQuery_GetJml(vprimary).HasValue = True Then
                If DTA_SO.ScalarQuery_GetJml(vprimary).Value > 0 Then
                    MsgBox("This data is has been used, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            DTA_SO.Dispose()

            Dim DTA_CodeSetting As New dseCustomerTableAdapters.tbltrans_soTableAdapter
            If DTA_CodeSetting.ScalarQuery_GetJml(vprimary).HasValue = True Then
                If DTA_CodeSetting.ScalarQuery_GetJml(vprimary).Value > 0 Then
                    MsgBox("This data is has been used, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            DTA_CodeSetting.Dispose()

            Dim DTA_CustComplain As New dseCustomerTableAdapters.tbltrans_soTableAdapter
            If DTA_CustComplain.ScalarQuery_GetJml(vprimary).HasValue = True Then
                If DTA_CustComplain.ScalarQuery_GetJml(vprimary).Value > 0 Then
                    MsgBox("This data is has been used, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            DTA_CustComplain.Dispose()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJmlBy_CustSegmentation(ByVal vPrimaryCustSegmentation As Double) As Integer
        Dim dta As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Try
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

    Public Function GetPrimary(ByVal ProductCode As String) As Double
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimary(ProductCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(ProductCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPhoneNo(ByVal vPrimaryCust As Double, ByVal vAddIdentifier As String) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customer_item_addressDataTable
        Dim row As dseCustomer.tblmaster_customer_item_addressRow

        Try
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

    Public Function GetNPWP(ByVal vPrimaryCust As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
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

    Public Function Get_City_ByPrimary(ByVal vPrimaryCust As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
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

    Public Function GetCustomerCode(ByVal vPrimary As Double) As String
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter

        Try
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

    Public Function IsiListBoxCustomer(ByVal lbo As CheckedListBox, ByVal TBL As dseCustomer.tblmaster_customer1DataTable) As Boolean
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customer1TableAdapter
        Dim row As dseCustomer.tblmaster_customer1Row

        Try
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

    Public Function IsiDGV(ByVal Criteria As String, _
     ByVal DGV As DataGridView, ByVal TBL As dseCustomer.DataTable1DataTable, _
     ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New dseCustomerTableAdapters.DataTable1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByAccountNo Then
                QueryResult = DTA.FillBy_AccNo(TBL, Criteria + "%")
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
                If TS.Ticks > 10000000 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Ticks / TimeSpan.TicksPerSecond) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, 0, MidpointRounding.AwayFromZero)) + " miliseconds"
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
            Return DTA.ScalarQuery_getTimeStamp(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal dgvAddress As DataGridView, _
    ByVal dgvItemCP As DataGridView, ByVal dgvItemBank As DataGridView) As Boolean
        Dim DTA As New dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim TBL As New dseCustomer.tblmaster_customerDataTable
        Dim row As dseCustomer.tblmaster_customerRow
        Dim DTA1 As New dseCustomerTableAdapters.tblmaster_customer_item_addressTableAdapter
        Dim TBL1 As New dseCustomer.tblmaster_customer_item_addressDataTable
        Dim row1 As dseCustomer.tblmaster_customer_item_addressRow
        Dim DTA2 As New dseCustomerTableAdapters.tblmaster_customer_itemcpTableAdapter
        Dim TBL2 As New dseCustomer.tblmaster_customer_itemcpDataTable
        Dim row2 As dseCustomer.tblmaster_customer_itemcpRow
        Dim DTA3 As New dseCustomerTableAdapters.tblmaster_customer_item_bankTableAdapter
        Dim TBL3 As New dseCustomer.tblmaster_customer_item_bankDataTable
        Dim row3 As dseCustomer.tblmaster_customer_item_bankRow

        Try
            DTA.FillBy_Primary(TBL, vPrimary)

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
                    Me.mPaymentTerm = .Payment_Terms
                    Me.mprimary = .Primary
                    Me.mprimaryCustomerGroup = .Primary_Customer_Group
                    Me.mprimaryCustomerSegmentation = .Primary_Customer_Segmentasi
                    Me.mprimaryCustomerType = .Primary_Customer_Type
                    Me.mprimarySalesArea = .Primary_Sales_Area
                    Me.mTaxAddressLine1 = .Tax_Add_Line1
                    Me.mTaxAddressLine2 = .Tax_Add_Line2
                    Me.mZipCode = .ZIP_Code

                    'salin data addres
                    dgvAddress.Rows.Clear()
                    dgvItemBank.Rows.Clear()
                    dgvItemCP.Rows.Clear()
                    DTA1.FillBy_PrimaryCustomer(TBL1, Me.mprimary)
                    Dim Obj2 As New ClsCekData
                    For Each row1 In TBL1.Rows
                        With row1
                            dgvAddress.Rows.Add("-", .Address_Code, .Line_1, .Line_2, .Line_3, .Line_4 _
                            , .Line_5, .Phone_No, .Fax_No, .Remark, .Identifier)
                        End With
                    Next
                    'salin data contact person
                    DTA2.FillBy_PrimaryCustomer(TBL2, Me.mprimary)
                    For Each row2 In TBL2.Rows
                        With row2
                            dgvItemCP.Rows.Add("-", .Name, .Position, .Phone_No, .Fax_No, _
                            ._E_mail_Address, .Remark, .Identifier)
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
        End Try
    End Function
End Class
