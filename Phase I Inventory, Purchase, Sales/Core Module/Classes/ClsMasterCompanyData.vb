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

    'Comment: Perubahan NPWP 16 Digit
    Private mNPWPLama As String = ""

    'Comment: Perubahan PPN 12%
    Private mTKU_ID As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    'Comment: Perubahan Perubahan PPN 12%
    Public Property TKU_ID() As String
        Get
            Return Me.mTKU_ID
        End Get
        Set(ByVal value As String)
            mTKU_ID = value
        End Set
    End Property

    'Comment: Perubahan NPWP 16 Digit
    Public Property NPWPLama() As String
        Get
            Return Me.mNPWPLama
        End Get
        Set(ByVal value As String)
            mNPWPLama = value
        End Set
    End Property

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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
                    'Comment: Perubahan NPWP 16 Digit
                    Me.mNPWPLama = .NPWP_Number_Old
                    'Comment: Perubahan PPN 12 %
                    Me.mTKU_ID = .TKU_ID
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
            'Comment: Perubahan PPN 12 %
            mTKU_ID = ""

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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.CompanyName, Me.AddLine1, Me.AddLine2, Me.AddLine3, Me.AddLine4, Me.AddLine5, Me.City, Me.NPWP, _
                          Me.PKPRegDate, Me.TaxAddLine1, Me.TaxAddLine2, vUserCode, Now, vUserCode, Now, Now, Me.mNPWPLama, Me.mTKU_ID) > 0 Then
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Update(Me.CompanyName, Me.AddLine1, Me.AddLine2, Me.AddLine3, Me.AddLine4, Me.AddLine5, Me.City, Me.NPWP, _
                          Me.PKPRegDate, Me.TaxAddLine1, Me.TaxAddLine2, vUserCode, Now, Now, Me.mNPWPLama, Me.mTKU_ID, CDec(Me.mPrimary), Me.mTimestamp) > 0 Then
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'simpan data Produk Code No
            For Each row In vDGVItem.Rows
                If row.Cells(Frm.ColAddressCode.Index).Value.ToString <> "" Then
                    If String.IsNullOrEmpty(row.Cells(Frm.Col_Line2.Index).Value.ToString) = True Then
                        row.Cells(Frm.Col_Line2.Index).Value = ""
                    End If
                    If String.IsNullOrEmpty(row.Cells(Frm.Col_Line3.Index).Value.ToString) = True Then
                        row.Cells(Frm.Col_Line3.Index).Value = ""
                    End If
                    If String.IsNullOrEmpty(row.Cells(Frm.Col_Line4.Index).Value.ToString) = True Then
                        row.Cells(Frm.Col_Line4.Index).Value = ""
                    End If
                    If String.IsNullOrEmpty(row.Cells(Frm.Col_Line5.Index).Value.ToString) = True Then
                        row.Cells(Frm.Col_Line5.Index).Value = ""
                    End If
                    If String.IsNullOrEmpty(row.Cells(Frm.ColRemarksAdd.Index).Value.ToString) = True Then
                        row.Cells(Frm.ColRemarksAdd.Index).Value = ""
                    End If

                    'simpan ke basis data item
                    DTA.Insert(CDec(vPrimaryTransaksi), row.Cells(Frm.ColAddressIdentifier.Index).Value, _
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Return row.Company_Name
            Else
                Return ""
            End If
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetAddress(ByVal vIdentifierAdd As String) As String
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_data_itemTableAdapter
        Dim TBL As New DseMasterCompanyData.tbltrans_company_data_itemDataTable
        Dim row As DseMasterCompanyData.tbltrans_company_data_itemRow
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)

            For Each row In TBL.Rows
                If row.Add_Identifier = vIdentifierAdd Then
                    Retval = row.Line_1
                    If row.Line_2 <> "" Then
                        If Retval = "" Then
                            Retval = row.Line_2
                        Else
                            Retval = Retval + vbCrLf + row.Line_2
                        End If
                    End If
                    If row.Line_3 <> "" Then
                        If Retval = "" Then
                            Retval = row.Line_3
                        Else
                            Retval = Retval + vbCrLf + row.Line_3
                        End If
                    End If
                    If row.Line_4 <> "" Then
                        If Retval = "" Then
                            Retval = row.Line_4
                        Else
                            Retval = Retval + vbCrLf + row.Line_4
                        End If
                    End If
                    If row.Line_5 <> "" Then
                        If Retval = "" Then
                            Retval = row.Line_5
                        Else
                            Retval = Retval + vbCrLf + row.Line_5
                        End If
                    End If

                    Return Retval
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan identifier addres pada baris pertama dari company data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetIdentifierShipToDefault() As String
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_data_itemTableAdapter
        Dim TBL As New DseMasterCompanyData.tbltrans_company_data_itemDataTable
        Dim row As DseMasterCompanyData.tbltrans_company_data_itemRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)

            If TBL.Rows.Count = 0 Then
                Return ""
            Else
                row = TBL.Rows(0)
                Return row.Add_Identifier
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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

    ''' <summary>
    ''' mengembalikan nilai true jika boleh hapus address yg diberikan identifiernya
    ''' </summary>
    ''' <param name="vIdentifier"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ApakahBolehHapusAddress(ByVal vIdentifier As String, ByVal vShowMsg As Boolean) As Boolean
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_poTableAdapter
        Dim TBL As New DseMasterCompanyData.tbltrans_poDataTable
        Dim row As DseMasterCompanyData.tbltrans_poRow
        Dim Dummy As String = ""
        Dim i As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'tidak boleh hapus jika ada transaksi po yg menggunakan data ini
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vIdentifier).HasValue = True Then
                If DTA.ScalarQuery_GetJml(vIdentifier).Value = 0 Then
                    Return True
                Else
                    If vShowMsg = True Then
                        DTA.FillBy_IdentifierAdd(TBL, vIdentifier)

                        While i <= 5 And i <= TBL.Rows.Count - 1
                            row = TBL.Rows(i)
                            If Dummy = "" Then
                                Dummy = row.No_Transaksi
                            Else
                                Dummy = Dummy + ", " + row.No_Transaksi
                            End If
                        End While
                        If TBL.Rows.Count > 5 Then
                            Dummy = Dummy + ", ect"
                        End If

                        MsgBox("This address referred by several PO i.e " + Dummy + ". You cannot delete this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If

                    Return False
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
                  ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
                  Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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

    Public Sub IsiCboAdd(ByVal vTBL As DseMasterCompanyData.tbltrans_company_data_itemDataTable, ByVal vCBO As ComboBox)
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_data_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(vTBL)
            vCBO.DataSource = vTBL.DefaultView
            vTBL.DefaultView.Sort = "AddressCode"
            vCBO.DisplayMember = "AddressCode"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    'Comment: Perubahan NPWP 16 Digit
    Public Function GetCompany_NPWP_Old() As String
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Dim TBL As New DseMasterCompanyData.tbltrans_company_dataDataTable
        Dim row As DseMasterCompanyData.tbltrans_company_dataRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Return row.NPWP_Number_Old
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

    'Comment: Perubahan PPN 12%
    Public Function GetTKU_ID() As String
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Dim TBL As New DseMasterCompanyData.tbltrans_company_dataDataTable
        Dim row As DseMasterCompanyData.tbltrans_company_dataRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Return row.TKU_ID
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

    Public Function GetTKU_ID(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterCompanyDataTableAdapters.tbltrans_company_dataTableAdapter
        Dim retval As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            retval = DTA.ScalarQuery_GetCompanyName(vPrimary)
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
