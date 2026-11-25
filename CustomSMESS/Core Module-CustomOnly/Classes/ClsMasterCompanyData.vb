Option Explicit On
Option Strict Off

Public Class ClsMasterCompanyData

    Private mPrimary As Double = -1
    Private mCompanyName As String = ""
    Private mAddLine1 As String = ""
    Private mAddLine2 As String = ""
    Private mAddLine3 As String = ""
    Private mAddLine4 As String = ""
    Private mAddLine5 As String = ""
    Private mCity As String = ""
    Private mNPWP As String = ""
    Private mPKPRegDate As Date = Now
    Private mTaxLine1 As String = ""
    Private mTaxLine2 As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Property CreatedBy() As String
        Get
            Return Me.mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return Me.mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Public Property ModifiedBy() As String
        Get
            Return Me.mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Public Property ModifiedDate() As Date
        Get
            Return Me.mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property TimeStamp() As Date
        Get
            Return Me.mTimestamp
        End Get
        Set(ByVal value As Date)
            Me.mTimestamp = value
        End Set
    End Property

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property CompanyName() As String
        Get
            Return mCompanyName
        End Get
        Set(ByVal value As String)
            mCompanyName = value
        End Set
    End Property

    Public Property AddLine1() As String
        Get
            Return mAddLine1
        End Get
        Set(ByVal value As String)
            mAddLine1 = value
        End Set
    End Property

    Public Property AddLine2() As String
        Get
            Return mAddLine2
        End Get
        Set(ByVal value As String)
            mAddLine2 = value
        End Set
    End Property

    Public Property AddLine3() As String
        Get
            Return mAddLine3
        End Get
        Set(ByVal value As String)
            mAddLine3 = value
        End Set
    End Property

    Public Property AddLine4() As String
        Get
            Return mAddLine4
        End Get
        Set(ByVal value As String)
            mAddLine4 = value
        End Set
    End Property

    Public Property AddLine5() As String
        Get
            Return mAddLine5
        End Get
        Set(ByVal value As String)
            mAddLine5 = value
        End Set
    End Property

    Public Property City() As String
        Get
            Return mCity
        End Get
        Set(ByVal value As String)
            mCity = value
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

    Public Property PKPRegDate() As Date
        Get
            Return mPKPRegDate
        End Get
        Set(ByVal value As Date)
            mPKPRegDate = value
        End Set
    End Property

    Public Property TaxAddLine1() As String
        Get
            Return mTaxLine1
        End Get
        Set(ByVal value As String)
            mTaxLine1 = value
        End Set
    End Property

    Public Property TaxAddLine2() As String
        Get
            Return mTaxLine2
        End Get
        Set(ByVal value As String)
            mTaxLine2 = value
        End Set
    End Property

    Public Function ApakahSudahAda(ByVal vCompanyName As String) As Boolean
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByCompanyName(vCompanyName).HasValue Then
                If DTA.ScalarQuery_GetJmlByCompanyName(vCompanyName).Value > 0 Then
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
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vDGVItem As DataGridView, ByVal FRM As FrmMasterCompanyData) As Boolean
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Dim TBL As New DseMasterCompanyData.tbltrans_company_dataDataTable
        Dim row As DseMasterCompanyData.tbltrans_company_dataRow
        Try
            DTA.Fill(TBL)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mCompanyName = .Company_Name
                    Me.mAddLine1 = .Add_Line_1
                    Me.mAddLine2 = .Add_Line_2
                    Me.mAddLine3 = .Add_Line_3
                    Me.mAddLine4 = .Add_Line_4
                    Me.mAddLine5 = .Add_Line_5
                    Me.mPKPRegDate = .PKP_Registration_Date
                    Me.mCity = .City
                    Me.mNPWP = .NPWP_Number
                    Me.mTaxLine1 = .Tax_Add_Line_1
                    Me.mTaxLine2 = .Tax_Add_Line_2
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                'salin data item ada 2 item 
                'CodeNo
                Me.IsiDGVItem(vDGVItem, FRM)

                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDGVItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimary = -1
            mCompanyName = ""
            mAddLine1 = ""
            mAddLine2 = ""
            mAddLine3 = ""
            mAddLine4 = ""
            mAddLine5 = ""
            mCity = ""
            mNPWP = ""
            mPKPRegDate = Now
            mTaxLine1 = ""
            mTaxLine2 = ""

            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimestamp = Today
            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetPrimary(ByVal vCompName As String) As Double
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Try
            If DTA.ScalarQuery_GetPrimary(vCompName).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vCompName).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Try
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanBaru(ByVal vDGVItem As DataGridView, _
      ByVal vUserCode As String, ByVal Frm As FrmMasterCompanyData) As Boolean
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Try
            If DTA.Insert(Me.CompanyName, Me.AddLine1, Me.AddLine2, Me.AddLine3, Me.AddLine4, Me.AddLine5, Me.City, Me.NPWP, _
                          Me.PKPRegDate, Me.TaxAddLine1, Me.TaxAddLine2, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                Me.mPrimary = Me.GetPrimary(Me.CompanyName)
                Me.mTimestamp = Me.GetTimeStamp(Me.mPrimary)
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'simpan data item
                Me.SimpanItem(vDGVItem, Me.mPrimary, Frm)

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vDGVItem As DataGridView, _
      ByVal vUserCode As String, ByVal Frm As FrmMasterCompanyData) As Boolean
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Try
            If DTA.Update(Me.CompanyName, Me.AddLine1, Me.AddLine2, Me.AddLine3, Me.AddLine4, Me.AddLine5, Me.City, Me.NPWP, _
                          Me.PKPRegDate, Me.TaxAddLine1, Me.TaxAddLine2, vUserCode, Now, Now, CDec(Me.mPrimary), Me.mTimestamp) > 0 Then
                Me.mTimestamp = CDate(DTA.ScalarQuery_GetTimestamp(Me.mPrimary))
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode
                'hapus data item lama
                Me.HapusItem(Me.mPrimary)
                'simpan data item baru
                Me.SimpanItem(vDGVItem, Me.mPrimary, Frm)
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''isi item
    Public Function IsiDGVItem(ByVal vDGVItem As DataGridView, _
         ByVal Frm As FrmMasterCompanyData) As Boolean
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_data_itemTableAdapter
        Dim TBL As New DseMasterCompanyData.tbltrans_company_data_itemDataTable
        Dim row As DseMasterCompanyData.tbltrans_company_data_itemRow
        Dim ObjCD As New ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim DefaultPrecision As Integer = 0
        Try
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()

            'isikan tabel
            If DTA.Fill(TBL) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColNo.Index).Value = ""
                        rowDGV.Cells(Frm.Col_Line1.Index).Value = .Line_1
                        rowDGV.Cells(Frm.Col_Line2.Index).Value = .Line_2
                        rowDGV.Cells(Frm.Col_Line3.Index).Value = .Line_3
                        rowDGV.Cells(Frm.Col_Line4.Index).Value = .Line_4
                        rowDGV.Cells(Frm.Col_Line5.Index).Value = .Line_5
                        rowDGV.Cells(Frm.ColRemarksAdd.Index).Value = .Remarks
                        rowDGV.Cells(Frm.ColAddressIdentifier.Index).Value = .Add_Identifier
                        rowDGV.Cells(Frm.ColAddressCode.Index).Value = .AddressCode
                    End With
                Next
            End If

            'set nomer
            ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
        End Try
    End Function

    Public Function SimpanItem(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
     ByVal Frm As FrmMasterCompanyData) As Boolean
        Dim row As DataGridViewRow
        Dim ObjCD As New ClsCekData
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_data_itemTableAdapter
        Dim DefaultPrecision As Integer = 0
        Try
            'simpan data Produk Code No
            For Each row In vDGVItem.Rows
                If row.Cells(Frm.ColAddressCode.Index).Value.ToString <> "" Then
                    If String.IsNullOrEmpty(row.Cells(Frm.Col_Line2.Index).Value.ToString) = True Then
                        row.Cells(Frm.Col_Line2.Index).Value = "-"
                    End If
                    If String.IsNullOrEmpty(row.Cells(Frm.Col_Line3.Index).Value.ToString) = True Then
                        row.Cells(Frm.Col_Line3.Index).Value = "-"
                    End If
                    If String.IsNullOrEmpty(row.Cells(Frm.Col_Line4.Index).Value.ToString) = True Then
                        row.Cells(Frm.Col_Line4.Index).Value = "-"
                    End If
                    If String.IsNullOrEmpty(row.Cells(Frm.Col_Line5.Index).Value.ToString) = True Then
                        row.Cells(Frm.Col_Line5.Index).Value = "-"
                    End If
                    If String.IsNullOrEmpty(row.Cells(Frm.ColRemarksAdd.Index).Value.ToString) = True Then
                        row.Cells(Frm.ColRemarksAdd.Index).Value = "-"
                    End If
                    'simpan ke basis data item
                    DTA.Insert(CDec(vPrimaryTransaksi), Me.Primary.ToString + "-" + row.Index.ToString, _
                    row.Cells(Frm.Col_Line1.Index).Value.ToString, row.Cells(Frm.Col_Line2.Index).Value.ToString, _
                    row.Cells(Frm.Col_Line3.Index).Value.ToString, row.Cells(Frm.Col_Line4.Index).Value.ToString, _
                    row.Cells(Frm.Col_Line5.Index).Value.ToString, row.Cells(Frm.ColRemarksAdd.Index).Value.ToString, _
                    row.Cells(Frm.ColAddressCode.Index).Value)
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_data_itemTableAdapter
        Try
            'hapus data item
            DTA.Delete(vPrimaryTransaksi)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''Pencarian data
    Public Function Pencarian(ByVal vDGV As DataGridView, ByVal FRM As FrmMasterCompanyData) As Boolean
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Dim TBL As New DseMasterCompanyData.tbltrans_company_dataDataTable
        Try
            DTA.Fill(TBL)
            If TBL.Rows.Count <= 0 Then
                Return False
            Else
                Me.CariDetail(vDGV, FRM)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetCompanyName(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCompanyName(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetCompanyName(vPrimary)
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

    Public Function GetCompanyName() As String
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Dim TBL As New DseMasterCompanyData.tbltrans_company_dataDataTable
        Dim row As DseMasterCompanyData.tbltrans_company_dataRow

        Try
            DTA.Fill(TBL)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Return row.Company_Name
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

    Public Function GetCompany_City() As String
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Dim TBL As New DseMasterCompanyData.tbltrans_company_dataDataTable
        Dim row As DseMasterCompanyData.tbltrans_company_dataRow

        Try
            DTA.Fill(TBL)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Return row.City
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

    Public Function GetCompany_NPWP() As String
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Dim TBL As New DseMasterCompanyData.tbltrans_company_dataDataTable
        Dim row As DseMasterCompanyData.tbltrans_company_dataRow

        Try
            DTA.Fill(TBL)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Return row.NPWP_Number
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

    Public Function GetCompany_TaxAdd1() As String
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Dim TBL As New DseMasterCompanyData.tbltrans_company_dataDataTable
        Dim row As DseMasterCompanyData.tbltrans_company_dataRow

        Try
            DTA.Fill(TBL)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Return row.Tax_Add_Line_1
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

    Public Function GetCompany_TaxAdd2() As String
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Dim TBL As New DseMasterCompanyData.tbltrans_company_dataDataTable
        Dim row As DseMasterCompanyData.tbltrans_company_dataRow

        Try
            DTA.Fill(TBL)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Return row.Tax_Add_Line_2
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

    Public Function GetCompany_TanggapPengukuhanPKP() As Date
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Dim TBL As New DseMasterCompanyData.tbltrans_company_dataDataTable
        Dim row As DseMasterCompanyData.tbltrans_company_dataRow

        Try
            DTA.Fill(TBL)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Return row.PKP_Registration_Date
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean) As Boolean
        Try
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mCompanyName) = True Then
                    MsgBox("Duplication in company name =" + Me.mCompanyName + ", please change your company name", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                'no po tidak boleh duplikat
                Dim NoLama As String = Me.GetCompanyName(Me.mPrimary)
                If Me.CompanyName <> NoLama Then
                    If Me.ApakahSudahAda(Me.mCompanyName) = True Then
                        MsgBox("Duplication in company name =" + Me.mCompanyName + ", please change your company name", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            'return value
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
                  ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
                  Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Try
            '2. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            If StrictEditting = True Then
                Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
                If CreatedBy <> vSalesStaff Then
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
