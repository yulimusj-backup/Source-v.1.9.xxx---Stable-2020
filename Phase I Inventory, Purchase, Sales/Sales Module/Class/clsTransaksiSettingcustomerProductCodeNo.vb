Option Explicit On
Option Strict Off

Public Class clsTransaksiSettingcustomerProductCodeNo
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mNoTransaksi As String = ""
    Private mNoUrut As Double = -1
    Private mTanggalTransaksi As Date = Now
    Private mPrimarycustomer As Double = -1
    Private mRemarks As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Property Primarycustomer() As Double
        Get
            Return Me.mPrimarycustomer
        End Get
        Set(ByVal value As Double)
            mPrimarycustomer = value
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

    Public Property TanggalTransaksi() As Date
        Get
            Return Me.mTanggalTransaksi
        End Get
        Set(ByVal value As Date)
            mTanggalTransaksi = value
        End Set
    End Property

    Public Property NoTransaksi() As String
        Get
            Return Me.mNoTransaksi
        End Get
        Set(ByVal value As String)
            mNoTransaksi = value
        End Set
    End Property

    Public Property NoUrut() As Double
        Get
            Return Me.mNoUrut
        End Get
        Set(ByVal value As Double)
            mNoUrut = value
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

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiCustomerCodeSettingTableAdapters.tbltrans_customerproductcodesettingTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            If DTA.ScalarQuery_getMaxNoUrut(TanggalStart, TanggalEnd).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_getMaxNoUrut(TanggalStart, TanggalEnd).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            'pastikan no urut terdiri atas 5 huruf
            While NoUrut.Length < 5
                NoUrut = "0" + NoUrut
            End While

            'bulan 
            Bulan = CStr(vTanggalTransaksi.Month)
            While Bulan.Length < 2
                Bulan = "0" + Bulan
            End While
            Return NoUrut + "/CUSTCODE/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiCustomerCodeSettingTableAdapters.tbltrans_customerproductcodesettingTableAdapter
        Try
            If DTA.ScalarQuery_getJmlNoTransaksi(vNoTransaksi).HasValue = True Then
                If DTA.ScalarQuery_getJmlNoTransaksi(vNoTransaksi).Value > 0 Then
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

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
    ByVal vDgvItem As DataGridView) As Boolean
        Dim DTA As New DseTransaksiCustomerCodeSettingTableAdapters.tbltrans_customerproductcodesettingTableAdapter
        Dim TBL As New DseTransaksiCustomerCodeSetting.tbltrans_customerproductcodesettingDataTable
        Dim row As DseTransaksiCustomerCodeSetting.tbltrans_customerproductcodesettingRow
        Dim ObjItem As New clsTransaksiSettingcustomerProductCodeNoItem

        Try
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mNoUrut = .Urutan
                    Me.mPrimary = .Primary
                    Me.mPrimarycustomer = .Primary_Customer
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                'salin data item 
                ObjItem.IsiDGV(vDgvItem, Me.mPrimary)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDgvItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit() As Boolean
        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus() As Boolean
        Try
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiCustomerCodeSettingTableAdapters.tbltrans_customerproductcodesettingTableAdapter
        Dim ObjItem As New clsTransaksiSettingcustomerProductCodeNoItem
        Dim SuppName As String
        Try
            'insert query
            If DTA.Insert(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 5)), Me.mTanggalTransaksi, Me.mPrimarycustomer, Me.mRemarks, vUserCode, _
            Now, vUserCode, Now, Now) > 0 Then
                'update data yang berubah setelah insert query
                Me.mPrimary = DTA.ScalarQuery_getPrimary(Me.mNoTransaksi).Value
                Me.mNoUrut = CInt(Me.mNoTransaksi.Substring(0, 5))
                Me.mTimestamp = DTA.ScalarQuery_getTimeStamp(Me.mPrimary).Value
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode
                If Me.mPrimarycustomer <> -1 Then
                    Dim ObjSupp As New Core_Module.clsMasterCustomer
                    SuppName = ObjSupp.GetCustName(Me.mPrimarycustomer)
                    ObjSupp = Nothing
                Else
                    SuppName = ""
                End If
                ObjItem.Simpan(vdgvItem, Me.mPrimary)

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaru(vdgvItem, vUserCode)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiCustomerCodeSettingTableAdapters.tbltrans_customerproductcodesettingTableAdapter
        Dim ObjItem As New clsTransaksiSettingcustomerProductCodeNoItem
        Dim SuppName As String
        Try
            If DTA.Update(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 5)), Me.mTanggalTransaksi, _
                Me.mPrimarycustomer, Me.mRemarks, vUserCode, Now, Now, Me.mPrimary, _
                Me.mTimestamp) > 0 Then
                'update data yang berubah setelah simpan perubahan
                Me.mNoUrut = Me.mNoTransaksi.Substring(0, 5)
                Me.mTimestamp = DTA.ScalarQuery_getTimeStamp(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode
                'hapus data item lama
                ObjItem.Hapus(Me.mPrimary)
                'simpan data item baru
                If Me.mPrimarycustomer <> -1 Then
                    Dim ObjSupp As New Core_Module.clsMasterCustomer
                    SuppName = ObjSupp.GetCustName(Me.mPrimarycustomer)
                    ObjSupp = Nothing
                Else
                    SuppName = ""
                End If
                ObjItem.Simpan(dgvItem, Me.mPrimary)
                'nilai balikan
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim Obj As New clsTransaksiSettingcustomerProductCodeNoItem
        Dim DTA As New DseTransaksiCustomerCodeSettingTableAdapters.tbltrans_customerproductcodesettingTableAdapter

        Try
            If Me.ApakahBolehHapus = True Then
                'hapus data item 
                Obj.Hapus(vPrimaryTransaksi)
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Obj = Nothing
        End Try
    End Function

    Public Function GetJmlUntukFormCari(ByVal Criteria As String, _
    ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiCustomerCodeSettingTableAdapters.tbltrans_customerproductcodesetting11TableAdapter
        Try
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                If DTA.ScalarQuery_getJmlNoTransaksiLike(Criteria + "%").HasValue = True Then
                    Return DTA.ScalarQuery_getJmlNoTransaksiLike(Criteria + "%")
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                If DTA.ScalarQuery_getJmlByTglDariSampai(vTanggalDari, vTanggalSampai).HasValue = True Then
                    Return DTA.ScalarQuery_getJmlByTglDariSampai(vTanggalDari, vTanggalSampai).Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                If DTA.ScalarQuery_getJmlByCustomerCodeLike(Criteria + "%").HasValue = True Then
                    Return DTA.ScalarQuery_getJmlByCustomerCodeLike(Criteria + "%")
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                If DTA.ScalarQuery_GetJmlByProductCodeLike(Criteria + "%").HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByProductCodeLike(Criteria + "%")
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                If DTA.ScalarQuery_getJmlByProductDescription(Criteria + "%").HasValue = True Then
                    Return DTA.ScalarQuery_getJmlByProductDescription(Criteria + "%")
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                If DTA.ScalarQuery_getJmlbyProductSpecification(Criteria + "%").HasValue = True Then
                    Return DTA.ScalarQuery_getJmlbyProductSpecification(Criteria + "%")
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

    Public Function IsiDGVFormCari(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiCustomerCodeSetting.tbltrans_customerproductcodesetting1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiCustomerCodeSettingTableAdapters.tbltrans_customerproductcodesetting1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TglDariSampai(TBL, vTanggalDari, vTanggalSampai)
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                QueryResult = DTA.FillBy_CustomerCodeLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_productDescription(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_ProductSpecification(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Transaction No"
                Dim ij As New DataGridViewCellStyle
                ij.Format = "dd MMM yyyy"
                DGV.Columns(2).DefaultCellStyle = ij
                DGV.Columns(2).HeaderText = "Transaction Date"
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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            Me.mCreatedBy = ""
            Me.mCreatedDate = Now
            Me.mModifiedBy = ""
            Me.mModifiedDate = Now
            Me.mNoTransaksi = ""
            Me.mNoUrut = -1
            Me.mPrimary = -1
            Me.mPrimarycustomer = -1
            Me.mRemarks = ""
            Me.mTanggalTransaksi = Now
            Me.mTimestamp = Now
            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiCustomerProductCodeNDesc(ByVal txtCodeNo As TextBox, ByVal txtDesc As TextBox, _
    ByVal PrimaryCust As Double, ByVal PrimaryProduct As Double)
        Dim DTA As New DseTransaksiCustomerCodeSettingTableAdapters.tblcustomercodesetting_viewTableAdapter
        Dim TBL As New DseTransaksiCustomerCodeSetting.tblcustomercodesetting_viewDataTable
        Dim row As DseTransaksiCustomerCodeSetting.tblcustomercodesetting_viewRow

        Try
            If DTA.Fill(TBL, PrimaryProduct, PrimaryCust) = 0 Then
                Dim ObjProduct As New Core_Module.clsMasterProduct
                txtCodeNo.Text = ObjProduct.GetProductCode(PrimaryProduct)
                txtDesc.Text = ObjProduct.GetDescription(PrimaryProduct)
                ObjProduct = Nothing
            Else
                row = TBL.Rows(0)
                txtCodeNo.Text = row._Customer_s_Code_No
                txtDesc.Text = row._Customer_s_Desc
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetPrimary(ByVal vCustomerProductCode As String, ByVal vPrimaryCust As Double) As Double
        Dim DTA As New DseTransaksiCustomerCodeSettingTableAdapters.tblcustomercodesetting_viewTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryByCustomerProductCode(vPrimaryCust, vCustomerProductCode).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryByCustomerProductCode(vPrimaryCust, vCustomerProductCode).Value
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
    EnumSearchCriteria, ByVal vPrimaryCust As Double) As Double
        Dim DTA As New DseTransaksiCustomerCodeSettingTableAdapters.tblcustomercodesetting_viewTableAdapter

        Try
            If vSearchBy = EnumSearchCriteria.ByProductCodeNo Then
                If DTA.ScalarQuery_GetJmlByProductCodeLIKE(vKriteria + "%", vPrimaryCust).HasValue Then
                    Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(vKriteria + "%", vPrimaryCust).Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByProductDesc Then
                If DTA.ScalarQuery_GetJmlByProductDescLIKE(vPrimaryCust, vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByProductDescLIKE(vPrimaryCust, vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByProductSpec Then
                If DTA.ScalarQuery_GetJmlByProductSpecLIKE(vPrimaryCust, vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByProductSpecLIKE(vPrimaryCust, vKriteria + "%").Value
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
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiCustomerCodeSetting.DataTable1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal vPrimaryCust _
    As Double) As Boolean
        Dim DTA As New DseTransaksiCustomerCodeSettingTableAdapters.DataTable1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, vPrimaryCust, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, vPrimaryCust, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_ProductSpecLIKE(TBL, vPrimaryCust, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
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

    Public Function GetJml(ByVal vProductCode As String, ByVal vPrimaryCust As Double) As Double
        Dim DTA As New DseTransaksiCustomerCodeSettingTableAdapters.tblcustomercodesetting_viewTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlByProductCode(vPrimaryCust, vProductCode).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByProductCode(vPrimaryCust, vProductCode).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
