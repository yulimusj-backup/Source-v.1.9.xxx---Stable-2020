Option Explicit On
Option Strict Off

Imports System.IO

Imports Microsoft.VisualBasic.FileIO
Public Class clsMasterProduct
    Dim mPrimary As Double = -1
    Dim mProductCode As String = ""
    Dim mPrimaryCategory As Double = -1
    Dim mSpecification As String = ""
    Dim mPrimaryUnit As Double = -1
    Dim mPrimaryMasterUnit As Double = -1
    Dim mQtyPerMasterUnit As Integer = -1
    Dim mPrimarySecondaryUnit As Double = -1
    Dim mBatchNoControl As Boolean = False
    Dim mLocationControl As Boolean = False
    Dim mUseSecondaryUnits As Boolean = False
    Dim mMinimumStock As Double = -1
    Dim mLeadTime As Integer = -1
    Dim mMinimumPurchaseQty As Double = -1
    Dim mDesc As String = ""
    Dim mPdoNoControlled As Boolean = False
    Dim mExpiryDateControlled As Boolean = False
    Dim mCategoryCodeBounded As Boolean = False
    Dim mCategoryDescBounded As Boolean = False
    Dim mDecimal As Integer = 0
    Dim mUseSizing As Integer = 0
    Dim mSize As Integer = 36
    Dim mBarcode As String = ""
    Dim mPrimaryAkunInventory As Double = -1
    Dim mPrimaryAkunWIP As Double = -1
    Dim mTipe As EnumTipeProduct = EnumTipeProduct.Goods

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Enum EnumTipeProduct As Integer
        Goods
        Services
    End Enum

    Public Property PrimaryAkunInventory() As Double
        Get
            Return Me.mPrimaryAkunInventory
        End Get
        Set(ByVal value As Double)
            mPrimaryAkunInventory = value
        End Set
    End Property
    Public Property PrimaryAkunWIP() As Double
        Get
            Return Me.mPrimaryAkunWIP
        End Get
        Set(ByVal value As Double)
            mPrimaryAkunWIP = value
        End Set
    End Property
    Public Property Barcode() As String
        Get
            Return Me.mBarcode
        End Get
        Set(ByVal value As String)
            mBarcode = value
        End Set
    End Property
    Public Property Size() As Integer
        Get
            Return Me.mSize
        End Get
        Set(ByVal value As Integer)
            Me.mSize = value
        End Set
    End Property
    Public Property UseSizing() As Boolean
        Get
            If Me.mUseSizing = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = True Then
                Me.mUseSizing = 1
            Else
                Me.mUseSizing = 0
            End If
        End Set
    End Property
    Public Property AngkaDecimal() As Integer
        Get
            Return Me.mDecimal
        End Get
        Set(ByVal value As Integer)
            mDecimal = value
        End Set
    End Property

    Public Property CategoryDescBounded() As Boolean
        Get
            Return Me.mCategoryDescBounded
        End Get
        Set(ByVal value As Boolean)
            Me.mCategoryDescBounded = value
        End Set
    End Property

    Public Property CategoryCodeBounded() As Boolean
        Get
            Return Me.mCategoryCodeBounded
        End Get
        Set(ByVal value As Boolean)
            mCategoryCodeBounded = value
        End Set
    End Property

    Public Property ExpiryDateControlled() As Boolean
        Get
            Return Me.mExpiryDateControlled
        End Get
        Set(ByVal value As Boolean)
            mExpiryDateControlled = value
        End Set
    End Property

    Public Property PdONumberControlled() As Boolean
        Get
            Return Me.mPdoNoControlled
        End Get
        Set(ByVal value As Boolean)
            mPdoNoControlled = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return Me.mDesc
        End Get
        Set(ByVal value As String)
            mDesc = value
        End Set
    End Property

    Public Property UseSecondaryUnits() As Boolean
        Get
            Return Me.mUseSecondaryUnits
        End Get
        Set(ByVal value As Boolean)
            mUseSecondaryUnits = value
        End Set
    End Property

    Public Property MinimumPurchaseQty() As Double
        Get
            Return Me.mMinimumPurchaseQty
        End Get
        Set(ByVal value As Double)
            mMinimumPurchaseQty = value
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

    Public Property MinimumStock() As Double
        Get
            Return Me.mMinimumStock
        End Get
        Set(ByVal value As Double)
            Me.mMinimumStock = value
        End Set
    End Property

    Public Property LocationControl() As Boolean
        Get
            Return Me.mLocationControl
        End Get
        Set(ByVal value As Boolean)
            mLocationControl = value
        End Set
    End Property

    Public Property BatchNoControl() As Boolean
        Get
            Return Me.mBatchNoControl
        End Get
        Set(ByVal value As Boolean)
            mBatchNoControl = value
        End Set
    End Property

    Public Property PrimarySecondaryUnit() As Double
        Get
            Return Me.mPrimarySecondaryUnit
        End Get
        Set(ByVal value As Double)
            mPrimarySecondaryUnit = value
        End Set
    End Property

    Public Property QtyPerMasterUnit() As Double
        Get
            Return Me.mQtyPerMasterUnit
        End Get
        Set(ByVal value As Double)
            mQtyPerMasterUnit = value
        End Set
    End Property

    Public Property PrimaryMasterUnit() As Double
        Get
            Return Me.mPrimaryMasterUnit
        End Get
        Set(ByVal value As Double)
            mPrimaryMasterUnit = value
        End Set
    End Property

    Public Property PrimaryUnit() As Double
        Get
            Return Me.mPrimaryUnit
        End Get
        Set(ByVal value As Double)
            mPrimaryUnit = value
        End Set
    End Property

    Public Property Specification() As String
        Get
            Return Me.mSpecification
        End Get
        Set(ByVal value As String)
            mSpecification = value
            If Me.mPrimaryCategory <> -1 And Me.mCategoryDescBounded = True Then
                Dim obj As New ClsMasterCategory
                Me.mDesc = obj.GetDesc(Me.mPrimaryCategory) + " " + Me.mSpecification
                obj = Nothing
            Else
                Me.mDesc = Me.mSpecification
            End If
        End Set
    End Property

    Public Property PrimaryCategory() As Double
        Get
            Return Me.mPrimaryCategory
        End Get
        Set(ByVal value As Double)
            mPrimaryCategory = value
            If Me.mPrimaryCategory <> -1 And Me.mCategoryDescBounded = True Then
                Dim obj As New ClsMasterCategory
                Me.mDesc = obj.GetDesc(Me.mPrimaryCategory) + " " + Me.mSpecification
                obj = Nothing
            Else
                Me.mDesc = Me.mSpecification
            End If
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

    Public Property ProductCode() As String
        Get
            Return Me.mProductCode
        End Get
        Set(ByVal value As String)
            mProductCode = value
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

    Public Function SimpanBaru(ByVal vUserCode As String, ByVal pbx As PictureBox, _
    ByVal dgvPrice As Windows.Forms.DataGridView) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter
        Dim row As DataGridViewRow
        Dim DTA1 As New dseMasterProductTableAdapters.tblmaster_product_pricingTableAdapter

        Try
            If pbx.ImageLocation <> "" Then
                '' Create a new stream to load this photo into
                Dim myStream As IO.FileStream = New FileStream(pbx.ImageLocation, FileMode.Open, _
                FileAccess.Read)
                '' Create a buffer to hold the stream of bytes
                Dim myImageBuffer(myStream.Length) As Byte
                '' Read the bytes from this stream and put it into the image buffer
                myStream.Read(myImageBuffer, 0, Convert.ToInt32(myStream.Length))
                '' Close the stream
                myStream.Close()

                If DTA.InsertQuery(Me.mProductCode, Me.mSpecification, Me.mPrimaryCategory, Me.mPrimaryUnit, Me.mPrimaryMasterUnit, _
                Me.mPrimarySecondaryUnit, Me.mQtyPerMasterUnit, Me.mBatchNoControl, Me.mLocationControl, _
                Me.mUseSecondaryUnits, myImageBuffer, vUserCode, Now, vUserCode, Now, Now, Me.mMinimumStock, _
                Me.mLeadTime, Me.mMinimumPurchaseQty, Me.mDesc, Me.mPdoNoControlled, Me.mExpiryDateControlled, _
                Me.mCategoryCodeBounded, Me.mCategoryDescBounded, Me.mDecimal, Me.mUseSizing, Me.mSize, Me.mTipe, Me.mPrimaryAkunInventory, _
                Me.mPrimaryAkunWIP, Me.mBarcode) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mProductCode, Me.mPrimaryCategory).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'input data harga
                    For Each row In dgvPrice.Rows
                        DTA1.Insert(Me.mPrimary, row.Cells.Item(1).Value, row.Cells(3).Value, row.Cells(4).Value)
                    Next

                    'return value
                    Return True
                Else
                    Return False
                End If
            Else
                If DTA.InsertQuery(Me.mProductCode, Me.mSpecification, Me.mPrimaryCategory, Me.mPrimaryUnit, Me.mPrimaryMasterUnit, _
                Me.mPrimarySecondaryUnit, Me.mQtyPerMasterUnit, Me.mBatchNoControl, Me.mLocationControl, _
                Me.mUseSecondaryUnits, Nothing, vUserCode, Now, vUserCode, Now, Now, Me.mMinimumStock, Me.mLeadTime, _
                Me.mMinimumPurchaseQty, Me.mDesc, Me.mPdoNoControlled, Me.mExpiryDateControlled, Me.mCategoryCodeBounded, _
                Me.mCategoryDescBounded, Me.mDecimal, Me.mUseSizing, Me.mSize, Me.mTipe, Me.mPrimaryAkunInventory, _
                Me.mPrimaryAkunWIP, Me.mBarcode) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mProductCode, Me.mPrimaryCategory).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'input data harga
                    For Each row In dgvPrice.Rows
                        DTA1.Insert(Me.mPrimary, row.Cells.Item(1).Value, row.Cells(3).Value, row.Cells(4).Value)
                    Next

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' fungsi ini mengembalikan nilai property yang dikontrol untuk suatu product
    ''' </summary>
    ''' <param name="vPrimaryProduct"></param>
    ''' <param name="Location"></param>
    ''' <param name="BatchNo"></param>
    ''' <param name="ExpiryDate"></param>
    ''' <param name="SecondaryUnit"></param>
    ''' <remarks></remarks>
    Public Sub GetProductProperties(ByVal vPrimaryProduct As Double, ByRef Location As Boolean, ByRef BatchNo As Boolean, _
    ByRef ExpiryDate As Boolean, ByRef SecondaryUnit As Boolean)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product3TableAdapter
        Dim TBL As New dseMasterProduct.tblmaster_product3DataTable
        Dim row As dseMasterProduct.tblmaster_product3Row

        Try
            If DTA.FillBy(TBL, vPrimaryProduct) > 0 Then
                row = TBL.Rows(0)
                If CBool(row.Batch_No_Controlled) = True Then
                    BatchNo = True
                Else
                    BatchNo = False
                End If
                If CBool(row.Location_Controlled) = True Then
                    Location = True
                Else
                    Location = False
                End If
                If CBool(row.Expiry_Date_Controlled) = True Then
                    ExpiryDate = True
                Else
                    ExpiryDate = False
                End If
                If CBool(row.Use_Secondary_Unit) = True Then
                    SecondaryUnit = True
                Else
                    SecondaryUnit = False
                End If
            Else
                Location = False
                BatchNo = False
                ExpiryDate = False
                SecondaryUnit = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Sub IsiCheckBoxControll(ByVal vPrimaryProduct As Double, ByVal chkLocation As CheckBox, ByVal chkBatchNo As CheckBox, ByVal chkExpiryDate As CheckBox, _
    ByVal chkSecondaryUnit As CheckBox, ByVal chkPdoControlled As CheckBox, ByVal TxtPrimaryUnit As TextBox, _
    ByVal TxtPrimarySecondaryUnits As TextBox)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product3TableAdapter
        Dim TBL As New dseMasterProduct.tblmaster_product3DataTable
        Dim row As dseMasterProduct.tblmaster_product3Row

        Try
            If DTA.FillBy(TBL, vPrimaryProduct) > 0 Then
                row = TBL.Rows(0)
                If CBool(row.Batch_No_Controlled) = True Then
                    chkBatchNo.Checked = True
                Else
                    chkBatchNo.Checked = False
                End If
                If CBool(row.Location_Controlled) = True Then
                    chkLocation.Checked = True
                Else
                    chkLocation.Checked = False
                End If
                If CBool(row.Expiry_Date_Controlled) = True Then
                    chkExpiryDate.Checked = True
                Else
                    chkExpiryDate.Checked = False
                End If
                If CBool(row.PdOControlled) = True Then
                    chkPdoControlled.Checked = True
                Else
                    chkPdoControlled.Checked = False
                End If
                If CBool(row.Use_Secondary_Unit) = True Then
                    chkSecondaryUnit.Checked = True
                Else
                    chkSecondaryUnit.Checked = False
                End If
                TxtPrimarySecondaryUnits.Text = row.Primary_Secondary_Unit
                TxtPrimaryUnit.Text = row.Primary_Unit
            Else
                chkLocation.Checked = False
                chkBatchNo.Checked = False
                chkExpiryDate.Checked = False
                chkSecondaryUnit.Checked = False
                chkPdoControlled.Checked = False
                TxtPrimarySecondaryUnits.Text = "-1"
                TxtPrimaryUnit.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function IsiListBoxProduct(ByVal lbo As CheckedListBox, ByVal TBL As dseMasterProduct.tblmaster_product2DataTable) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product2TableAdapter

        Try
            DTA.Fill(TBL)
            lbo.DataSource = TBL.DefaultView
            lbo.DisplayMember = "Combination"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiCheckBoxControll(ByVal vPrimaryProduct As Double, ByVal chkLocation As CheckBox, _
    ByVal chkBatchNo As CheckBox, ByVal chkExpiryDate As CheckBox, ByVal TxtPrimaryUnit As TextBox)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product3TableAdapter
        Dim TBL As New dseMasterProduct.tblmaster_product3DataTable
        Dim row As dseMasterProduct.tblmaster_product3Row

        Try
            If DTA.FillBy(TBL, vPrimaryProduct) > 0 Then
                row = TBL.Rows(0)
                If CBool(row.Batch_No_Controlled) = True Then
                    chkBatchNo.Checked = True
                Else
                    chkBatchNo.Checked = False
                End If
                If CBool(row.Location_Controlled) = True Then
                    chkLocation.Checked = True
                Else
                    chkLocation.Checked = False
                End If
                If CBool(row.Expiry_Date_Controlled) = True Then
                    chkExpiryDate.Checked = True
                Else
                    chkExpiryDate.Checked = False
                End If
                TxtPrimaryUnit.Text = row.Primary_Unit
            Else
                chkLocation.Checked = False
                chkBatchNo.Checked = False
                chkExpiryDate.Checked = False
                TxtPrimaryUnit.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       EnumSearchCriteria) As Double
        Dim DTA As New dseMasterProductTableAdapters.DataTable1TableAdapter

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
            ElseIf vSearchBy = EnumSearchCriteria.ByProductSpec Then
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

    Public Function ApakahKodeTerakhirMenggunakanBarcode() As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter
        Dim vPrimary As Double = -1

        Try
            If DTA.ScalarQuery_GetPrimaryMax.HasValue = False Then
                Return True
            Else
                vPrimary = DTA.ScalarQuery_GetPrimaryMax.Value
                If String.IsNullOrEmpty(DTA.ScalarQuery_GetBarcode(vPrimary)) = True Then
                    Return False
                Else
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.dispose()
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vUserCode As String, ByVal pbx As PictureBox, ByVal dgvPrice As DataGridView) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter
        Dim DTA1 As New dseMasterProductTableAdapters.tblmaster_product_pricingTableAdapter

        Dim row As DataGridViewRow
        Try
            'cek apakah ada perubahan pada batch no controled, expiry date controled, location controlled,
            'PdO controlled, use secondary units
            If pbx.ImageLocation <> "" Then 'ada perubahan data gambar
                '' Create a new stream to load this photo into
                Dim myStream As IO.FileStream = New FileStream(pbx.ImageLocation, FileMode.Open, _
                FileAccess.Read)
                '' Create a buffer to hold the stream of bytes
                Dim myImageBuffer(myStream.Length) As Byte
                '' Read the bytes from this stream and put it into the image buffer
                myStream.Read(myImageBuffer, 0, Convert.ToInt32(myStream.Length))
                '' Close the stream
                myStream.Close()

                If DTA.Update(Me.mProductCode, Me.mSpecification, Me.mPrimaryCategory, Me.mPrimaryUnit, _
                Me.mPrimaryMasterUnit, Me.mPrimarySecondaryUnit, Me.mQtyPerMasterUnit, Me.mBatchNoControl, _
                Me.mLocationControl, Me.mUseSecondaryUnits, myImageBuffer, vUserCode, Now, Now, Me.mMinimumStock, _
                Me.mLeadTime, Me.mMinimumPurchaseQty, Me.mDesc, Me.mPdoNoControlled, Me.mExpiryDateControlled, _
                Me.mCategoryCodeBounded, Me.mCategoryDescBounded, Me.mDecimal, Me.mUseSizing, Me.mSize, Me.mTipe, _
                Me.mPrimaryAkunInventory, Me.mPrimaryAkunWIP, Me.mBarcode, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data lama
                    DTA1.Delete(Me.mPrimary)
                    'input data harga
                    For Each row In dgvPrice.Rows
                        DTA1.Insert(Me.mPrimary, row.Cells.Item(1).Value, row.Cells(3).Value, row.Cells(4).Value)
                    Next

                    Return True
                Else
                    Return False
                End If
            Else
                If IsNothing(pbx.Image) = True Then
                    'tidak ada data gambar yang perlu disimpan
                    If DTA.Update(Me.mProductCode, Me.mSpecification, Me.mPrimaryCategory, Me.mPrimaryUnit, _
                    Me.mPrimaryMasterUnit, Me.mPrimarySecondaryUnit, Me.mQtyPerMasterUnit, Me.mBatchNoControl, _
                    Me.mLocationControl, Me.mUseSecondaryUnits, Nothing, vUserCode, Now, Now, Me.mMinimumStock, _
                    Me.mLeadTime, Me.mMinimumPurchaseQty, Me.mDesc, Me.mPdoNoControlled, Me.mExpiryDateControlled, _
                    Me.mCategoryCodeBounded, Me.mCategoryDescBounded, Me.mDecimal, Me.mUseSizing, Me.mSize, Me.mTipe, _
                    Me.mPrimaryAkunInventory, Me.mPrimaryAkunWIP, Me.mBarcode, Me.mPrimary, Me.mTimestamp) > 0 Then
                        Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                        Me.mModifiedDate = Me.mTimestamp
                        Me.mModifiedBy = vUserCode

                        'hapus data lama
                        DTA1.Delete(Me.mPrimary)
                        'input data harga
                        For Each row In dgvPrice.Rows
                            DTA1.Insert(Me.mPrimary, row.Cells.Item(1).Value, row.Cells(3).Value, row.Cells(4).Value)
                        Next

                        Return True
                    Else
                        Return False
                    End If
                Else
                    'data gambar tidak mengalami perubahan
                    If DTA.UpdateQuery(Me.mProductCode, Me.mSpecification, Me.mPrimaryCategory, Me.mPrimaryUnit, _
                    Me.mPrimaryMasterUnit, Me.mPrimarySecondaryUnit, Me.mQtyPerMasterUnit, Me.mBatchNoControl, _
                    Me.mLocationControl, Me.mUseSecondaryUnits, vUserCode, Now, Now, Me.mMinimumStock, _
                    Me.mLeadTime, Me.mMinimumPurchaseQty, Me.mDesc, Me.mPdoNoControlled, Me.mExpiryDateControlled, _
                    Me.mCategoryCodeBounded, Me.mCategoryDescBounded, Me.mDecimal, Me.mUseSizing, _
                    Me.mSize, Me.mTipe, Me.mPrimaryAkunInventory, Me.mPrimaryAkunWIP, Me.mBarcode, Me.mPrimary, Me.mTimestamp) > 0 Then
                        Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                        Me.mModifiedDate = Me.mTimestamp
                        Me.mModifiedBy = vUserCode

                        'hapus data lama
                        DTA1.Delete(Me.mPrimary)
                        'input data harga
                        For i As Integer = 0 To dgvPrice.Rows.Count - 1
                            DTA1.Insert(Me.mPrimary, dgvPrice.Rows(i).Cells(1).Value, dgvPrice.Rows(i).Cells(3).Value, _
                            dgvPrice.Rows(i).Cells(4).Value)
                        Next

                        Return True
                    Else
                        Return False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If Me.ApakahBolehHapus(vPrimary) = True Then
                DTA.Delete(vPrimary)    'data pricing otomatis di delete oleh server
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

    Public Function ApakahSudahAda(ByVal vProductCode As String) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_GetJml2(vProductCode).HasValue Then
                If DTA.ScalarQuery_GetJml2(vProductCode).Value > 0 Then
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

    Public Function ApakahSudahAdaBarcode(ByVal vBarcode As String) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlByBarcode(vBarcode).HasValue Then
                If DTA.ScalarQuery_GetJmlByBarcode(vBarcode).Value > 0 Then
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

    Public Function IsiListBatchNo(ByVal vPrimaryProduct As Double, ByVal LS As List(Of String), _
    ByVal LED As List(Of Date)) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqty1TableAdapter
        Dim TBL As New dseMasterProduct.tblmaster_productqty1DataTable
        Dim row As dseMasterProduct.tblmaster_productqty1Row

        Try
            'isi database
            DTA.Fill(TBL, vPrimaryProduct)
            'kosongkan list
            LS.Clear()
            LED.Clear()
            'isi berdasarkan yang ada didalam basis data
            For Each row In TBL.Rows
                LS.Add(row.Lot_No)
                LED.Add(row.Expiry_Date)
            Next
            'return value
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function ApakahLocationControlled(ByVal vPrimary As Double) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_IsLocationControlled(vPrimary).HasValue = True Then
                Return CBool(DTA.ScalarQuery_IsLocationControlled(vPrimary).Value)
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan nilai true jika sudah pernah digunakan pada PO, SO, Stock In Out, atau bundling
    ''' </summary>
    ''' <param name="vPrimaryProduct"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ApakahSudahPernahDigunakan(ByVal vPrimaryProduct As Double) As Boolean
        Try
            'cek, jika sudah ada penggunaan code no di:
            '1. po
            'Dim DTA As New dseMasterProductTableAdapters.tbltrans_po_itemTableAdapter
            'If DTA.ScalarQuery_GetJmlPOItemByProduct(vPrimaryProduct).Value > 0 Then
            ' DTA.Dispose()
            'Return True
            'End If
            'DTA.Dispose()
            '2. so
            'Dim DTA1 As New dseMasterProductTableAdapters.tbltrans_so_itemTableAdapter
            'If DTA1.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
            'DTA1.Dispose()
            'Return True
            'End If
            'DTA1.Dispose()
            '3. stock in out
            Dim DTA2 As New dseMasterProductTableAdapters.tbltransaksi_stockinout_itemTableAdapter
            If DTA2.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                DTA2.Dispose()
                Return True
            End If
            DTA2.Dispose()
            '4. stock bundling
            Dim DTA3 As New dseMasterProductTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
            If DTA3.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                DTA3.Dispose()
                Return True
            End If
            DTA3.Dispose()
            '5. product qty
            Dim DTA4 As New dseMasterProductTableAdapters.tblmaster_productqty5TableAdapter
            If DTA4.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).HasValue = True Then
                If DTA4.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                    DTA4.Dispose()
                    Return True
                End If
            End If
            DTA4.Dispose()

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBatchNoControlled(ByVal vPrimary As Double) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_IsBatchNoControlled(vPrimary).HasValue = True Then
                Return CBool(DTA.ScalarQuery_IsBatchNoControlled(vPrimary).Value)
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlPOSByPrimary(ByVal vPrimary As Double) As Decimal
        Dim DTA As New dseMasterProductTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByPrimaryProduct(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimaryProduct(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimaryProduct As Double) As Boolean
        Try
            'tidak boleh dihapus bila sudah ada yang merefer:
            '1. Transaksi inventory
            'stock in out
            Dim DTA2 As New dseMasterProductTableAdapters.tbltransaksi_stockinout_itemTableAdapter
            If DTA2.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                DTA2.Dispose()
                Return False
            End If
            DTA2.Dispose()
            'stock transfer
            Dim DTA5 As New dseMasterProductTableAdapters.tbltransaksi_stocktransfer_itemTableAdapter
            If DTA5.ScalarQuery_GetJml(vPrimaryProduct).Value > 0 Then
                DTA5.Dispose()
                Return False
            End If
            DTA5.Dispose()
            'stock bundling item in/out
            Dim DTA3 As New dseMasterProductTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
            If DTA3.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                DTA3.Dispose()
                Return False
            End If
            Dim DTA16 As New dseMasterProductTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
            If DTA16.ScalarQuery_GetJml(vPrimaryProduct).Value > 0 Then
                DTA16.Dispose()
                Return False
            End If
            DTA3.Dispose()
            'stock tag item
            Dim DTA6 As New dseMasterProductTableAdapters.tbltrans_stock_taking_tag_itemTableAdapter
            If DTA6.ScalarQuery_GEtJml(vPrimaryProduct).Value > 0 Then
                DTA6.Dispose()
                Return False
            End If
            DTA6.Dispose()
            'product qty
            Dim DTA4 As New dseMasterProductTableAdapters.tblmaster_productqty5TableAdapter
            If DTA4.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).HasValue = True Then
                If DTA4.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                    DTA4.Dispose()
                    Return False
                End If
            End If
            DTA4.Dispose()
            '2. Transaksi sales
            'so
            Dim DTA7 As New dseMasterProductTableAdapters.tbltrans_so_itemTableAdapter
            If DTA7.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value > 0 Then
                DTA7.Dispose()
                Return False
            End If
            'quotation
            Dim DTA8 As New dseMasterProductTableAdapters.tbltrans_quotation_itemTableAdapter
            If DTA8.ScalarQuery_GetJml(vPrimaryProduct).Value > 0 Then
                DTA8.Dispose()
                Return False
            End If
            'customer product code no
            Dim DTA11 As New dseMasterProductTableAdapters.tbltrans_customerproductcodestting_itemTableAdapter
            If DTA11.ScalarQuery_GetJml(vPrimaryProduct).Value > 0 Then
                DTA11.Dispose()
                Return False
            End If
            '3. Transaksi puchase
            'historical purchase
            Dim DTA12 As New dseMasterProductTableAdapters.tbltrans_historicalpurchase_itemTableAdapter
            If DTA12.ScalarQuery_GetJml(vPrimaryProduct).Value > 0 Then
                DTA12.Dispose()
                Return False
            End If
            'po
            Dim DTA13 As New dseMasterProductTableAdapters.tbltrans_po_item1TableAdapter
            If DTA13.ScalarQuery_GetJml(vPrimaryProduct).Value > 0 Then
                DTA13.Dispose()
                Return False
            End If
            'monthly purchase
            Dim DTA15 As New dseMasterProductTableAdapters.tbltrans_purchase_closing_itemTableAdapter
            If DTA15.ScalarQuery_GetJml(vPrimaryProduct).Value > 0 Then
                DTA15.Dispose()
                Return False
            End If

            'Transaksi POS
            If Me.GetJmlPOSByPrimary(vPrimaryProduct) > 0 Then
                MsgBox("Data cannot be deleted because there are other transaction that refering this data", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'kalo sampe sini maka boleh hapus
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub IsiCboLocation(ByVal vPrimaryProduct As Double, ByVal vCbo As ComboBox, _
    ByVal TBL As dseMasterProduct.stock_location_viewDataTable, ByVal vBatchNo As String)
        Dim DTA As New dseMasterProductTableAdapters.stock_location_viewTableAdapter

        Try
            vCbo.DataSource = Nothing
            vCbo.Items.Clear()
            vCbo.Refresh()
            If DTA.Fill(TBL, vPrimaryProduct, vBatchNo) > 0 Then
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiCboSecondaryUnit(ByVal vPrimaryProduct As Double, ByVal vCbo As ComboBox, _
        ByVal TBL As dseMasterProduct.tblmaster_productqty3DataTable, ByVal vBatchNo As String)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqty3TableAdapter

        Try
            vCbo.DataSource = Nothing
            vCbo.Items.Clear()
            vCbo.Refresh()

            If DTA.Fill(TBL, vPrimaryProduct, vBatchNo) > 0 Then
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiCboSecondaryUnit(ByVal vPrimaryProduct As Double, ByVal vCbo As ComboBox, _
    ByVal TBL As dseMasterProduct.tblmaster_productqty2DataTable, ByVal vBatchNo As String, _
    ByVal vPrimaryLocation As Double)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqty2TableAdapter

        Try
            vCbo.DataSource = Nothing
            vCbo.Items.Clear()
            vCbo.Refresh()

            If DTA.Fill(TBL, vPrimaryProduct, vBatchNo, vPrimaryLocation) > 0 Then
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLUnitPrice(ByVal vPrimaryProduct As Double, ByVal vTBL As dseMasterProduct.DataTable3DataTable)
        Dim DTA As New dseMasterProductTableAdapters.DataTable3TableAdapter

        Try
            DTA.FillBy_PrimaryProduct(vTBL, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiCboBatchNo(ByVal vPrimaryProduct As Double, ByVal vCbo As ComboBox, _
    ByVal TBL As dseMasterProduct.tblmaster_productqty1DataTable)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqty1TableAdapter

        Try
            vCbo.DataSource = Nothing
            vCbo.Refresh()
            vCbo.Items.Clear()
            If DTA.Fill(TBL, vPrimaryProduct) > 0 Then
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiDGVBacthNo(ByVal vPrimaryProduct As Double, ByVal vDGV As DataGridView, _
    ByVal TBL As dseMasterProduct.tblmaster_productqty1DataTable)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqty1TableAdapter

        Try
            If DTA.Fill(TBL, vPrimaryProduct) > 0 Then
                vDGV.DataSource = TBL.DefaultView
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByRef TBL As dseMasterProduct.tblmaster_product_view0DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByRef DTA As dseMasterProductTableAdapters.tblmaster_product_view0TableAdapter) As Boolean
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.Fill(TBL)
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DESCLike(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_ProductSpecLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByBarcode Then
                QueryResult = DTA.FillBy_BarcodeLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Product code"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
            End If


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
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    ''' <summary>
    ''' mengisikan dgv, tetapi hanya untuk item - item yang location controlled
    ''' </summary>
    ''' <param name="Criteria"></param>
    ''' <param name="DGV"></param>
    ''' <param name="TBL"></param>
    ''' <param name="SearchCriteria"></param>
    ''' <param name="tsl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsiDGVLocationControlledOnly(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As dseMasterProduct.tblmaster_product_view0DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product_view0TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE_LocationControlled(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE_LocationControlled(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_SpecLIKE_LocationControlled(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByBarcode Then
                QueryResult = DTA.FillBy_SpecLIKE_LocationControlled(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Product code"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(9).Visible = True
            End If


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

    Public Function CariDetail(ByVal Primary As Double, ByVal dgvPricing As DataGridView, ByVal pbx As PictureBox) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product3TableAdapter
        Dim TBL As New dseMasterProduct.tblmaster_product3DataTable
        Dim row As dseMasterProduct.tblmaster_product3Row
        Dim DTA1 As New dseMasterProductTableAdapters.tblmaster_product_pricingTableAdapter
        Dim TBL1 As New dseMasterProduct.tblmaster_product_pricingDataTable
        Dim row1 As dseMasterProduct.tblmaster_product_pricingRow

        Try
            DTA.FillBy(TBL, Primary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = row.Primary
                    Me.mBatchNoControl = .Batch_No_Controlled
                    Me.mLocationControl = .Location_Controlled
                    Me.mExpiryDateControlled = .Expiry_Date_Controlled
                    Me.mCategoryCodeBounded = .Category_Code_Bounded
                    Me.mCategoryDescBounded = .Category_Desc_Bounded
                    Me.mPrimaryCategory = .Primary_Category
                    Me.mPrimaryMasterUnit = .Primary_Master_Unit
                    Me.mPrimarySecondaryUnit = .Primary_Secondary_Unit
                    Me.mPrimaryUnit = .Primary_Unit
                    Me.mProductCode = .Product_Code
                    Me.mQtyPerMasterUnit = .Qty_Per_Master
                    Me.mSpecification = .Specification
                    Me.mUseSecondaryUnits = .Use_Secondary_Unit
                    Me.mLeadTime = .Lead_Time
                    Me.mMinimumPurchaseQty = .Minimum_Purchase_Qty
                    Me.mMinimumStock = .Minimum_Stock
                    Me.mDesc = .Description
                    Me.mPdoNoControlled = .PdOControlled
                    Me.mDecimal = ._Decimal

                    Me.mUseSizing = .Use_Sizing
                    Me.mSize = .Size
                    Me.mBarcode = .Barcode
                    Me.mTipe = .Type
                    Me.mPrimaryAkunInventory = .Primary_Inventory_Account
                    Me.mPrimaryAkunWIP = .Primary_WIP_Account

                    'salin data gambar
                    If .IsPictureNull = False Then
                        Dim myByteArray() As Byte = .Picture

                        If myByteArray.Length > 0 Then
                            '' Create a new MemoryStream and write all the information from
                            '' the byte array into the stream
                            Dim myStream As MemoryStream = New MemoryStream(myByteArray, True)
                            myStream.Write(myByteArray, 0, myByteArray.Length)

                            '' Use the MemoryStream to create the new BitMap object
                            Dim FinalImage As Bitmap = New Bitmap(myStream)

                            pbx.Image = FinalImage
                            '' Close the stream
                            myStream.Close()
                        End If
                    Else
                        pbx.Image = Nothing
                    End If

                    'salin data pricing
                    dgvPricing.Rows.Clear()
                    DTA1.FillBy(TBL1, Me.mPrimary)
                    Dim Obj1 As New clsMasterCurrency
                    Dim Obj2 As New ClsCekData
                    For Each row1 In TBL1.Rows
                        With row1
                            dgvPricing.Rows.Add("-", .Primary_Currency, Obj1.GetCurrencyCode(.Primary_Currency), _
                            .Unit_Price, .Discount)
                        End With
                    Next
                    Obj2.SetNumber(dgvPricing)
                    Obj1 = Nothing
                    Obj2 = Nothing

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
        End Try
    End Function

    Public Function GetProductCode(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCode(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetCode(vPrimary)
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

    Public Function GetSecondaryUnitCode(ByVal vPrimaryProduct As Double) As String
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetSecondaryUnitsCode(vPrimaryProduct)) = False Then
                Return DTA.ScalarQuery_GetSecondaryUnitsCode(vPrimaryProduct)
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

    Public Function GetUnitCode(ByVal vPrimaryProduct As String) As String
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetUnitCode(vPrimaryProduct)) = False Then
                Return DTA.ScalarQuery_GetUnitCode(vPrimaryProduct)
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

    Public Function GetStockAwal(ByVal vPrimary As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_GetStockAwal(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetStockAwal(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' fungsi ini mengembalikan jumlah product yang "bounded by category" untuk category tertentu
    ''' </summary>
    ''' <param name="vPrimaryCategory"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJml_ProductBoundedByCategory(ByVal vPrimaryCategory As Double) As Long
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlBounded(vPrimaryCategory).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlBounded(vPrimaryCategory).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            Return DTA.ScalarQuery_GetTimestamp(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDescription(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDesc(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetDesc(vPrimary)
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

    Public Function ApakahUseSecondary(ByVal vPrimaryProduct As Double) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_IsUseSecondary(vPrimaryProduct).HasValue = True Then
                Return CBool(DTA.ScalarQuery_IsUseSecondary(vPrimaryProduct).Value)
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetAngkaDecimal(ByVal vPrimaryProduct As Double) As Integer
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_GetDecimal(vPrimaryProduct).HasValue Then
                Return DTA.ScalarQuery_GetDecimal(vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vPrimaryCategory As Double) As Integer
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlByCategory(vPrimaryCategory).HasValue Then
                Return DTA.ScalarQuery_GetJmlByCategory(vPrimaryCategory).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vProductCode As String) As Long
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_GetJml2(vProductCode).HasValue = True Then
                Return DTA.ScalarQuery_GetJml2(vProductCode).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetExpiryDate(ByVal vPrimaryProduct As Double, ByVal vBatchNo As String) As Date
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Try
            If DTA.ScalarQuery_GetExpiryDate(vPrimaryProduct, vBatchNo).HasValue = True Then
                Return DTA.ScalarQuery_GetExpiryDate(vPrimaryProduct, vBatchNo).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByPrimary(ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByPrimaryProduct(vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_ByUnit(ByVal vPrimaryUnit As Double) As Integer
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter
        Dim dummy As Integer = 0

        Try
            If DTA.ScalarQuery_GetJml_ByPrimaryMasterUnit(vPrimaryUnit).HasValue Then
                dummy += DTA.ScalarQuery_GetJml_ByPrimaryMasterUnit(vPrimaryUnit).Value
            End If
            If DTA.ScalarQuery_GetJmlByPrimarySecondaryUnit(vPrimaryUnit).HasValue Then
                dummy += DTA.ScalarQuery_GetJmlByPrimarySecondaryUnit(vPrimaryUnit).Value
            End If
            If DTA.ScalarQuery_GetJmlByPrimaryUnit(vPrimaryUnit).HasValue Then
                dummy += DTA.ScalarQuery_GetJmlByPrimaryUnit(vPrimaryUnit).Value
            End If

            Return dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_ByStorageArea(ByVal vPrimaryStorageArea As Double) As Integer
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlByPrimaryStorageArea(vPrimaryStorageArea).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimaryStorageArea(vPrimaryStorageArea).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlBy_PrimaryCurrency(ByVal vPrimaryCurrency As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product_pricingTableAdapter

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

    Public Function GetSpecification(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_getSpecification(vPrimary)) = False Then
                Return DTA.ScalarQuery_getSpecification(vPrimary)
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

    Public Function GetDiscount(ByVal vPrimaryCurrency As Double, ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product_pricingTableAdapter

        Try
            If DTA.ScalarQuery_GetDiscount(vPrimaryProduct, vPrimaryCurrency).HasValue = True Then
                Return DTA.ScalarQuery_GetDiscount(vPrimaryProduct, vPrimaryCurrency).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetUnitPrice(ByVal vPrimaryCurrency As Double, ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product_pricingTableAdapter

        Try
            If DTA.ScalarQuery_GetUnitPrice(vPrimaryProduct, vPrimaryCurrency).HasValue = True Then
                Return DTA.ScalarQuery_GetUnitPrice(vPrimaryProduct, vPrimaryCurrency).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanQtyBaru(ByVal vJmlBaru As Decimal, ByVal vPrimaryProduct As Double, ByVal DTA1 As dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, _
    Optional ByVal vSecondaryQty As Decimal = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal vBatchNo As String = "", _
    Optional ByVal vExpiryDate As Date = #1/1/1950#) As Boolean

        Try
            If DTA1.Insert(vPrimaryProduct, vJmlBaru, vSecondaryQty, vBatchNo, vExpiryDate, vPrimaryLocation) > 0 Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJml2(ByVal vPrimaryProduct As Double, ByVal DTA1 As dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, _
    Optional ByVal vBatchNo As String = "") As Decimal
        Try
            If DTA1.ScalarQuery_GetJml(vPrimaryProduct, vBatchNo, vPrimaryLocation, vSecondaryQty).HasValue = True Then
                Return DTA1.ScalarQuery_GetJml(vPrimaryProduct, vBatchNo, vPrimaryLocation, vSecondaryQty).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJml3(ByVal vPrimaryProduct As Double, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, _
    Optional ByVal vBatchNo As String = "") As Decimal
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Try
            Return Me.GetJml2(vPrimaryProduct, DTA, vSecondaryQty, vPrimaryLocation, _
            vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml4(ByVal vPrimaryProduct As Double, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vBatchNo As String = "") As Decimal
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Try
            Return DTA.ScalarQuery_GetJml2(vPrimaryProduct, vBatchNo, vSecondaryQty).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml5(ByVal vPrimaryCodeNo As Double, ByVal vSecQty As Double) As Decimal
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Try
            Return DTA.ScalarQuery_GetJml3(vPrimaryCodeNo, vSecQty).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan jumlah product yang location controlled
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJml_LocationControlled() As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlLocationControlled.HasValue = True Then
                Return DTA.ScalarQuery_GetJmlLocationControlled.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub SetBatchNumberControlled(ByVal vPrimaryProduct As Double, ByVal vControlled As Boolean)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.UpdateQuery_SetBatchNumberControlled(vControlled, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub SetExpiryDateControlled(ByVal vPrimaryProduct As Double, ByVal vControlled As Boolean)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.UpdateQuery_SetEDControlled(vControlled, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub SetLocationControl(ByVal vPrimaryProduct As Double, ByVal vControlled As Boolean)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.UpdateQuery_SetLocationControlled(vControlled, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub DeleteProductQty(ByVal vPrimaryProduct As Double)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Try
            DTA.DeleteQuery(vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub InsertProductQty(ByVal vPrimaryProduct As Double, ByVal vDGV As DataGridView, ByVal vColBatch As Integer, _
    ByVal vColED As Integer, ByVal vColQty As Integer, ByVal vColPrimaryLocation As Integer, ByVal vQtyNoControl As Double)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim row As DataGridViewRow

        Try
            If vDGV.RowCount = 0 Then
                If vQtyNoControl > 0 Then
                    DTA.Insert(vPrimaryProduct, vQtyNoControl, -1, _
                    "", #1/1/1950#, -1)
                End If
            Else
                'secondary qty =-1
                For Each row In vDGV.Rows
                    With row
                        DTA.Insert(vPrimaryProduct, .Cells(vColQty).Value, -1, _
                        .Cells(vColBatch).Value, .Cells(vColED).Value, .Cells(vColPrimaryLocation).Value)
                    End With
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub SetTanggalMinimum(ByVal vPrimaryProduct As Double, ByVal vTanggalMinimum As Date)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.UpdateQuery_SetTanggalMinimum(vTanggalMinimum, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function SetJml(ByVal vJmlBaru As Decimal, ByVal vPrimaryProduct As Double, ByVal DTA1 As dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, _
    Optional ByVal vBatchNo As String = "") As Boolean
        Try
            If DTA1.UpdateQuery_SetJml(vJmlBaru, vBatchNo, vPrimaryLocation, vPrimaryProduct, vSecondaryQty) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryProduct As Double, ByVal DTA1 As dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal vBatchNo As String = "") As Boolean
        Try
            If DTA1.Delete(vPrimaryProduct, vBatchNo, vPrimaryLocation, vSecondaryQty) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetPrimaryUnit(ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryUnit(vPrimaryProduct).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryUnit(vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCategory(ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryCategory(vPrimaryProduct).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCategory(vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vProductCode As String, ByVal vPrimaryCategory As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimary(vProductCode, vPrimaryCategory).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vProductCode, vPrimaryCategory).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryBrand(ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryBrand(vPrimaryProduct).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryBrand(vPrimaryProduct).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vProductCode As String) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimary2(vProductCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary2(vProductCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryByBarcode(ByVal vBarcode As String) As Double
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter
        Try
            If DTA.ScalarQuery_GetPrimaryByBarcode(vBarcode).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryByBarcode(vBarcode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahExpiryDateControlled(ByVal vPrimaryProduct As Double) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_IsiExpiryDateControlled(vPrimaryProduct).HasValue = True Then
                Return CBool(DTA.ScalarQuery_IsiExpiryDateControlled(vPrimaryProduct).Value)
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Sub IsiCboBatchNo_UntukUpdateBatch(ByVal vPrimaryProduct As Double, ByVal vCbo As ComboBox, _
 ByVal TBL As dseMasterProduct.tblmaster_productqty1DataTable)
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_productqty1TableAdapter
        Dim DTAIn As New dseMasterProductTableAdapters.DataTable5TableAdapter
        Dim TBLIn As New dseMasterProduct.DataTable5DataTable
        Dim rowIn As dseMasterProduct.DataTable5Row
        Dim row As dseMasterProduct.tblmaster_productqty1Row
        Dim i As Integer = 0

        Try
            vCbo.DataSource = Nothing
            vCbo.Refresh()
            vCbo.Items.Clear()
            DTA.Fill(TBL, vPrimaryProduct)

            'tambahkan dengan semua stock in yang belum ada dalam tabel diatas
            DTAIn.Fill(TBLIn, vPrimaryProduct)
            For Each row In TBL.Rows
                'buang baris2 yang ada di tblIn, dimana batch numbernya sudah ada di dalam TBL
                i = 0
                While i < TBLIn.Rows.Count
                    rowIn = TBLIn.Rows(i)
                    If rowIn.Batch_No.ToLower = row.Lot_No.ToLower Then
                        TBLIn.Rows.Remove(rowIn)
                    Else
                        i += 1
                    End If
                End While
            Next
            For Each rowIn In TBLIn.Rows
                row = TBL.Newtblmaster_productqty1Row
                row.Lot_No = rowIn.Batch_No
                row.Expiry_Date = rowIn.Expiry_Date
                row.Qty = 0
                TBL.Rows.Add(row)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAIn.Dispose()
            TBLIn.Dispose()
        End Try
    End Sub

    Public Function GetBarcode(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product4TableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetBarcode(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetBarcode(vPrimary)
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
