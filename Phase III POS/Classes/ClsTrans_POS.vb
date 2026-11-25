Option Explicit On
Option Strict Off
Imports Core_Module

Public Class ClsTrans_POS

    Private mPrimary As Double = -1
    Private mNoTrans As String = ""
    Private mTanggal As Date = Now
    Private mUrutan As Double = 0
    Private mPrimarySession As Double = -1
    Private mPrimaryMember As Double = -1
    Private mPrimaryTransServer As Double = -1
    Private mPrimaryDisc As Double = -1
    Private mPrimaryAccountDiscCost As Double = -1
    Private mDiscStr As String = ""
    Private mUseTax As Integer = 0
    Private mTaxPersen As Decimal = 0
    Private mIsSinglePayment As Integer = 0
    Private mSubTotal As Decimal = 0
    Private mDiscount As Decimal = 0
    Private mTotal As Decimal = 0
    Private mTaxAmount As Decimal = 0
    Private mRounding As Decimal = 0
    Private mBankCharge As Decimal = 0
    Private mGrandTotal As Decimal = 0
    Private mTotalPayment As Decimal = 0
    Private mChange As Decimal = 0
    Private mIsVoid As Integer = 0
    Private mRemarks As String = ""
    Private mPrimaryCurr As Double = -1
    Private mDefPrec As Double = 0

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Property PrimaryCurr() As Double
        Get
            Return mPrimaryCurr
        End Get
        Set(ByVal value As Double)
            mPrimaryCurr = value
        End Set
    End Property

    Public Property DefPrec() As Double
        Get
            Return mDefPrec
        End Get
        Set(ByVal value As Double)
            mDefPrec = value
        End Set
    End Property

    Public Property PrimaryMember() As Double
        Get
            Return mPrimaryMember
        End Get
        Set(ByVal value As Double)
            mPrimaryMember = value
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

    Public Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Public Property IsVoid() As Integer
        Get
            Return mIsVoid
        End Get
        Set(ByVal value As Integer)
            mIsVoid = value
        End Set
    End Property

    Public Property Change() As Decimal
        Get
            Return mChange
        End Get
        Set(ByVal value As Decimal)
            mChange = value
        End Set
    End Property

    Public Property TotalPayment() As Decimal
        Get
            Return mTotalPayment
        End Get
        Set(ByVal value As Decimal)
            mTotalPayment = value
        End Set
    End Property

    Public Property GrandTotal() As Decimal
        Get
            Return mGrandTotal
        End Get
        Set(ByVal value As Decimal)
            mGrandTotal = value
        End Set
    End Property

    Public Property BankCharge() As Decimal
        Get
            Return mBankCharge
        End Get
        Set(ByVal value As Decimal)
            mBankCharge = value
        End Set
    End Property

    Public Property Rounding() As Decimal
        Get
            Return mRounding
        End Get
        Set(ByVal value As Decimal)
            mRounding = value
        End Set
    End Property

    Public Property TaxAmount() As Decimal
        Get
            Return mTaxAmount
        End Get
        Set(ByVal value As Decimal)
            mTaxAmount = value
        End Set
    End Property

    Public Property Total() As Decimal
        Get
            Return mTotal
        End Get
        Set(ByVal value As Decimal)
            mTotal = value
        End Set
    End Property

    Public Property Subtotal() As Double
        Get
            Return mSubTotal
        End Get
        Set(ByVal value As Double)
            mSubTotal = value
        End Set
    End Property

    Public Property Discount() As Decimal
        Get
            Return mDiscount
        End Get
        Set(ByVal value As Decimal)
            mDiscount = value
        End Set
    End Property

    Public Property IsSinglePayment() As Integer
        Get
            Return mIsSinglePayment
        End Get
        Set(ByVal value As Integer)
            mIsSinglePayment = value
        End Set
    End Property

    Public Property TaxPersen() As Decimal
        Get
            Return mTaxPersen
        End Get
        Set(ByVal value As Decimal)
            mTaxPersen = value
        End Set
    End Property

    Public Property UseTax() As Integer
        Get
            Return mUseTax
        End Get
        Set(ByVal value As Integer)
            mUseTax = value
        End Set
    End Property

    Public Property DiscStr() As String
        Get
            Return mDiscStr
        End Get
        Set(ByVal value As String)
            mDiscStr = value
        End Set
    End Property

    Public Property PrimarySession() As Double
        Get
            Return mPrimarySession
        End Get
        Set(ByVal value As Double)
            mPrimarySession = value
        End Set
    End Property

    Public Property PrimaryTransServer() As Double
        Get
            Return mPrimaryTransServer
        End Get
        Set(ByVal value As Double)
            mPrimaryTransServer = value
        End Set
    End Property

    Public Property PrimaryDisc() As Double
        Get
            Return mPrimaryDisc
        End Get
        Set(ByVal value As Double)
            mPrimaryDisc = value
        End Set
    End Property

    Public Property PrimaryAccountDiscCost() As Double
        Get
            Return mPrimaryAccountDiscCost
        End Get
        Set(ByVal value As Double)
            mPrimaryAccountDiscCost = value
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

    Public Property NoTrans() As String
        Get
            Return mNoTrans
        End Get
        Set(ByVal value As String)
            mNoTrans = value
        End Set
    End Property

    Public Property Tanggal() As Date
        Get
            Return mTanggal
        End Get
        Set(ByVal value As Date)
            mTanggal = value
        End Set
    End Property

    Public Property Urutan() As Double
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Double)
            mUrutan = value
        End Set
    End Property

    Public Enum EnumModeCetak As Integer
        'ada 3 macam format cetakan yang bisa digenerate oleh pos
        'khusus untuk kepeng, guna format cetak continuous
        StghLegal
        A5
        Continuous
    End Enum

    Public Sub Cetak(ByVal vPrimaryTransaksi As Double, ByVal vModeCetak As EnumModeCetak)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetConstring() As String
        Try
            Return "server=" + My.Settings.server + ";user id=" + My.Settings.user_id + ";Password=" + My.Settings.password + _
            ";persist security info=True;database=" + My.Settings.database_name + ";port=" + My.Settings.port_number + ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=localhost;user id=root;Password=123;persist security info=True;database=core_db;port=3307"
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date, ByVal vPrimarySession As Double) As String
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Dim ObjSession As New ClsTrans_Session
        Dim ObjBranch As New ClsMast_Cabang
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData
        Dim BranchStr As String = "-"
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetMaxUrutan(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetMaxUrutan(TanggalEnd, TanggalStart).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            'pastikan no urut terdiri atas 6 huruf
            While NoUrut.Length < 6
                NoUrut = "0" + NoUrut
            End While

            'bulan 
            Bulan = CStr(vTanggalTransaksi.Month)
            While Bulan.Length < 2
                Bulan = "0" + Bulan
            End While

            'Branch 
            BranchStr = ObjBranch.GetCode(ObjSession.GetPrimaryBranch(vPrimarySession))

            Return NoUrut + "/" + "POS" + "/" + BranchStr + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            ObjSession = Nothing
            ObjBranch = Nothing
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByNotransPOS(vNoTransaksi).HasValue = True Then
                If DTA.ScalarQuery_GetJmlByNotransPOS(vNoTransaksi).Value > 0 Then
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

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vTanggal As Date, _
    ByVal vPrimarySession As Double) As Boolean
        Try
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mNoTrans) = True Then
                    Me.mNoTrans = Me.GetNoBaru(vTanggal, vPrimarySession)
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetPrimary(ByVal vNoPOS As String) As Double
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimary(vNoPOS).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vNoPOS).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoPOS(ByVal vPrimary As Double) As String
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoPOS(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetNoPOS(vPrimary)
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

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBY(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetCreatedBY(vPrimary)
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

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
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

    Public Overridable Function SimpanBaru(ByVal vTanggal As Date, ByVal vPrimarySession As Double, _
    ByVal vdgvItem As DataGridView, ByVal vUserCode As String, ByVal i As FrmTrans_POS) As Boolean
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Dim ObjItem As New ClsTrans_POS_Item
        Try
            If Me.ApakahDataBenar(True, vTanggal, vPrimarySession) = True Then
                Me.mUrutan = CInt(Me.mNoTrans.Substring(0, 6))
                If DTA.Insert(Me.NoTrans, Me.Tanggal, Me.Urutan, Me.PrimarySession, Me.PrimaryMember, _
                Me.PrimaryTransServer, Me.PrimaryDisc, Me.PrimaryAccountDiscCost, Me.DiscStr, Me.UseTax, _
                Me.TaxPersen, Me.IsSinglePayment, Me.Subtotal, Me.Discount, Me.Total, Me.TaxAmount, _
                Me.Rounding, Me.BankCharge, Me.GrandTotal, Me.TotalPayment, Me.Change, Me.IsVoid, vUserCode, Now, _
                vUserCode, Now, Now, Me.Remarks, Me.PrimaryCurr, Me.DefPrec) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTrans).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item baru
                    ObjItem.SimpanItem(Me.mPrimary, vdgvItem, i)
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTrans + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTrans = Me.GetNoBaru(Me.mTanggal, Me.mPrimarySession)
                Return Me.SimpanBaru(vTanggal, vPrimarySession, vdgvItem, vUserCode, i)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function UpdateDataHeader(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            Me.mUrutan = Me.mNoTrans.Substring(0, 6)
            DTA.UpdateQuery_SimpanPerubahan(Me.mTanggal, Me.mUrutan, Me.mPrimarySession, Me.mPrimaryMember, _
                Me.mPrimaryTransServer, Me.mPrimaryDisc, Me.mPrimaryAccountDiscCost, Me.DiscStr, Me.mUseTax, _
                Me.mTaxPersen, Me.mIsSinglePayment, Me.mSubTotal, Me.mDiscount, Me.mTotal, Me.mTaxAmount, Me.mRounding, _
                Me.mBankCharge, Me.mGrandTotal, Me.mTotalPayment, Me.mChange, Me.mIsVoid, vUserCode, Now, Now, Me.mRemarks, _
                Me.mPrimaryCurr, Me.mDefPrec, Me.Primary, Me.TimeStamp, Me.mNoTrans)
            Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
            Me.mModifiedDate = Me.mTimestamp
            Me.mModifiedBy = vUserCode
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function UpdateIsVoid(ByVal vPrimary As Double, ByVal vIsVoid As Integer) As Boolean
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.UpdateQuery_IsVoid(vIsVoid, vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal vPrimarySession As Double, ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTrans_POS.DataTable3DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTrans_POSTableAdapters.DataTable3TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Dim TglAwal As String
        Dim TglAkhir As String
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_POSCodeLIKE(TBL, Criteria + "%", vPrimarySession)
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                TglAkhir = Format(vTanggalSampai, " dd MMM yyyy") + " " + "23:59:59"
                TglAwal = Format(vTanggalDari, "dd MMM yyyy") + " " + "00:00:00"
                QueryResult = DTA.FillBy_TanggalDariSampai(TBL, TglAkhir, TglAwal, vPrimarySession)
            ElseIf SearchCriteria = EnumSearchCriteria.BySessionID Then
                QueryResult = DTA.FillBy_SessionIDLIKE(TBL, Criteria + "%", vPrimarySession)
            ElseIf SearchCriteria = EnumSearchCriteria.ByMemberCode Then
                QueryResult = DTA.FillBy_MemberCodeLIKE(TBL, Criteria + "%", vPrimarySession)
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%", vPrimarySession)
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "ID Number"
                DGV.Columns(2).HeaderText = "Date"
                DGV.Columns(3).HeaderText = "Session ID"
                DGV.Columns(4).HeaderText = "Member Code"
                DGV.Columns(5).HeaderText = "Remarks"
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
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
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml(ByVal vPrimarySession As Double, ByVal Criteria As String, _
    ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTrans_POSTableAdapters.DataTable3TableAdapter
        Try
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByPOSCodeLIKE(Criteria + "%", vPrimarySession).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.ScalarQuery_GetJmlByTanggalDariSampai(vTanggalSampai, vTanggalDari, vPrimarySession).Value
            ElseIf SearchCriteria = EnumSearchCriteria.BySessionID Then
                Return DTA.ScalarQuery_GetJmlBySessionIDLIKE(Criteria + "%", vPrimarySession).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByMemberCode Then
                Return DTA.ScalarQuery_GetJmlByMemberCodeLIKE(Criteria + "%", vPrimarySession).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                Return DTA.ScalarQuery_GetJmlBYRemarksLIKE(Criteria + "%", vPrimarySession).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function CariDetailHeaderPOS(ByVal vPrimaryTransaksi As Double, _
    ByVal vDgvItem As DataGridView, ByVal vObjCD As ClsCekData, ByVal i As FrmTrans_POS) As Boolean
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Dim TBL As New DseTrans_POS.tbltrans_pos_cabangDataTable
        Dim row As DseTrans_POS.tbltrans_pos_cabangRow
        Dim ObjItem As New ClsTrans_POS_Item
        Try
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)
                With row
                    mPrimary = .primary
                    mNoTrans = .no_trans
                    mTanggal = ._date
                    mUrutan = .urutan
                    mPrimarySession = .primary_session
                    mPrimaryMember = .primary_member
                    mPrimaryTransServer = .primary_transaksi_server
                    mPrimaryDisc = .primary_discount
                    mPrimaryAccountDiscCost = .primary_account_discount_cost
                    mDiscStr = .discount_string
                    mUseTax = .use_tax
                    mTaxPersen = .tax_percentage
                    mIsSinglePayment = .is_single_payment
                    mSubTotal = .sub_total
                    mDiscount = .discount
                    mTotal = .total
                    mTaxAmount = .tax_amount
                    mRounding = .rounding
                    mBankCharge = .bank_charge
                    mGrandTotal = .grand_total
                    mTotalPayment = .total_payment
                    mChange = .change
                    mIsVoid = .is_void
                    mRemarks = .remarks
                    mPrimaryCurr = .primary_curr
                    mDefPrec = .Default_Precision

                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With

                'salin data item 
                ObjItem.IsiDGV(vPrimaryTransaksi, vDgvItem, i, vObjCD)
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

    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.Delete(vPrimary)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryMAX() As Double
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimaryMaximum.HasValue Then
                Return DTA.ScalarQuery_GetPrimaryMaximum.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlPayment(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_item_paymentTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlPaymentByPrimaryTrans(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetJmlPaymentByPrimaryTrans(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function UpdatePrimaryMember(ByVal vPrimary As Double, ByVal vPrimaryMember As Double) As Boolean
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.UpdateQuery_PrimaryMember(vPrimaryMember, vPrimary)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub ResetObject(ByVal vDGVItem As DataGridView)
        Try
            mPrimary = -1
            mNoTrans = ""
            mTanggal = Now
            mUrutan = 0
            mPrimarySession = -1
            mPrimaryMember = -1
            mPrimaryTransServer = -1
            mPrimaryDisc = -1
            mPrimaryAccountDiscCost = -1
            mDiscStr = ""
            mUseTax = 0
            mTaxPersen = 0
            mIsSinglePayment = 0
            mSubTotal = 0
            mDiscount = 0
            mTotal = 0
            mTaxAmount = 0
            mRounding = 0
            mBankCharge = 0
            mGrandTotal = 0
            mTotalPayment = 0
            mChange = 0
            mIsVoid = 0
            mRemarks = ""
            mPrimaryCurr = -1
            mDefPrec = 0
            '-----
            mCreatedBy = ""
            mCreatedDate = Now
            mModifiedBy = ""
            mModifiedDate = Now
            mTimestamp = Now
            '-----
            vDGVItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function CariDetail(ByVal vPrimary As Double, _
    ByVal vDGVItem As DataGridView, ByVal vObjCD As ClsCekData, _
    ByVal i As FrmTrans_POS) As Boolean
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Dim TBL As New DseTrans_POS.tbltrans_pos_cabangDataTable
        Dim row As DseTrans_POS.tbltrans_pos_cabangRow
        Dim ObjItem As New ClsTrans_POS_Item
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .primary
                    Me.mNoTrans = .no_trans
                    Me.mTanggal = ._date
                    Me.mUrutan = .urutan
                    Me.mPrimarySession = .primary_session
                    Me.mPrimaryMember = .primary_member
                    Me.mPrimaryTransServer = .primary_transaksi_server
                    Me.mPrimaryDisc = .primary_discount
                    Me.mPrimaryAccountDiscCost = .primary_account_discount_cost
                    Me.mDiscStr = .discount_string
                    Me.mUseTax = .use_tax
                    Me.mTaxPersen = .tax_percentage
                    Me.mIsSinglePayment = .is_single_payment
                    Me.mSubTotal = .sub_total
                    Me.mDiscount = .discount
                    Me.mTotal = .total
                    Me.mTaxAmount = .tax_amount
                    Me.mRounding = .rounding
                    Me.mBankCharge = .bank_charge
                    Me.mGrandTotal = .grand_total
                    Me.mTotalPayment = .total_payment
                    Me.mChange = .change
                    Me.mIsVoid = .is_void
                    Me.mRemarks = .remarks
                    Me.mPrimaryCurr = .primary_curr
                    Me.mDefPrec = .Default_Precision

                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With
            End If

            ObjItem.IsiDGV(vPrimary, vDGVItem, i, vObjCD)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function GetFungsiListButtonNext(ByVal vPrimarySession As Double, ByVal vTBLNext As DseTrans_POS.TBLNextDataTable) As Boolean
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Dim TBL As New DseTrans_POS.tbltrans_pos_cabangDataTable
        Dim Row As DseTrans_POS.tbltrans_pos_cabangRow
        Dim RowNext As DseTrans_POS.TBLNextRow
        Dim ListPrimary As New List(Of Double)
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.FillBy_PrimarySessionNTotalPaymentNol(TBL, vPrimarySession)

            If vTBLNext.Rows.Count > 0 Then
                If TBL.Rows.Count > 0 Then
                    'List Primary Mana aja yang udah ada
                    For Each RowNext In vTBLNext.Rows
                        If ListPrimary.Contains(RowNext.PrimaryPOS) = False Then
                            ListPrimary.Add(RowNext.PrimaryPOS)
                        End If
                    Next
                    For Each Row In TBL.Rows
                        If ListPrimary.Contains(Row.primary) = False Then
                            ListPrimary.Add(Row.primary)
                            RowNext = vTBLNext.NewRow
                            With Row
                                RowNext.CheckPosisi = 0
                                RowNext.PrimaryPOS = .primary
                                RowNext.NoPOS = .no_trans
                            End With
                            vTBLNext.AddTBLNextRow(RowNext)
                        End If
                    Next
                    Return True
                Else
                    Return False
                End If
            Else
                If TBL.Rows.Count > 0 Then
                    For Each Row In TBL.Rows
                        RowNext = vTBLNext.NewRow
                        With Row
                            RowNext.CheckPosisi = 0
                            RowNext.PrimaryPOS = .primary
                            RowNext.NoPOS = .no_trans
                        End With
                        vTBLNext.AddTBLNextRow(RowNext)
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
            TBL.Dispose()
            ListPrimary = Nothing
        End Try
    End Function

    Public Function GetIsVoid(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetIsViod(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetIsViod(vPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Coding coba-coba 
    Public Sub CetakCobaCoba(ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer)
        Dim Report As New CrystalReport1
        Dim DS As New DseTrans_POS
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Dim TBL As DseTrans_POS.tbltrans_pos_cabangDataTable = DS.tbltrans_pos_cabang
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            DTA.Fill(TBL)

            Dim ij As New System.Drawing.Printing.PaperSource
            ij.SourceName = "85x55"
            Report.SetDataSource(DS)
            Report.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Portrait
            Report.PrintOptions.PrinterName = "EPSON LX-300+ /II"
            Report.PrintOptions.CustomPaperSource = ij

            'Report.SetParameterValue(Report.Parameter_TanggalSampai.ParameterFieldName, vEndDate.Date)
            'Report.PrintOptions.PaperSize.DefaultPaperSize = CrystalDecisions.Shared.PaperSize.DefaultPaperSize

            vCrv.ReportSource = Report
            vCrv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
        End Try
    End Sub

    Public Sub CetakPOSUkuranStngahLetter(ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vUserCode As String, ByVal vPrimary As Double, Optional ByVal LangsungCetak As Boolean = False)
        Dim Report As New CrpCetakPOSSetengahLetter
        Dim DS As New DseTrans_POS_Cetak
        Dim TBL As DseTrans_POS_Cetak.Tabel_CetakDataTable = DS.Tabel_Cetak
        Dim Row As DseTrans_POS_Cetak.Tabel_CetakRow
        Dim ObjProd As New clsMasterProduct
        Dim ObjUnit As New ClsMasterUnit
        Dim ObjPay As New ClsMast_PaymentOption
        Dim DTAHeader As New DseTrans_POS_CetakTableAdapters.tbltrans_pos_cabangTableAdapter
        Dim DTAItem As New DseTrans_POS_CetakTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Dim DTAPay As New DseTrans_POS_CetakTableAdapters.tbltrans_pos_cabang_item_paymentTableAdapter
        Dim TBLHeader As New DseTrans_POS_Cetak.tbltrans_pos_cabangDataTable
        Dim TBLItem As New DseTrans_POS_Cetak.tbltrans_pos_cabang_itemDataTable
        Dim TBLPay As New DseTrans_POS_Cetak.tbltrans_pos_cabang_item_paymentDataTable
        Dim RowHeader As DseTrans_POS_Cetak.tbltrans_pos_cabangRow
        Dim RowItem As DseTrans_POS_Cetak.tbltrans_pos_cabang_itemRow
        Dim RowPay As DseTrans_POS_Cetak.tbltrans_pos_cabang_item_paymentRow
        Dim IndexPayment As Integer = 0
        Dim JumlahTarikTunai As Decimal = 0
        Dim SubTotalStr As String = ""
        Dim SubTotalDec As String = ""
        Dim ObjCD As New ClsCekData

        Try
            DTAHeader.Connection.ConnectionString = Me.GetConstring
            DTAItem.Connection.ConnectionString = Me.GetConstring
            DTAPay.Connection.ConnectionString = Me.GetConstring

            DTAHeader.FillBy_Primary(TBLHeader, vPrimary)
            DTAItem.FillBy_Primary(TBLItem, vPrimary)
            DTAPay.FillBy_Primary(TBLPay, vPrimary)

            'Hitung dulu payment-nya 
            RowHeader = TBLHeader.Rows(0)
            If RowHeader.is_single_payment = 0 Then
                RowPay = TBLPay.Rows(0)
                If RowPay.primary_payment = "-1" Then
                    RowHeader.BayarDesc1 = "Tunai"
                    RowHeader.BayarTotal1 = RowPay.amount
                Else
                    RowHeader.BayarDesc1 = ObjPay.GetDesc(RowPay.primary_payment)
                    RowHeader.BayarTotal1 = RowPay.amount
                End If
            Else
                For Each RowPay In TBLPay.Rows
                    If IndexPayment <= 4 Then
                        If ObjPay.GetTypePayment(RowPay.primary_payment) = 2 Then
                            RowHeader.CashWithdrawal += RowPay.amount
                        Else
                            If IndexPayment = 0 Then
                                If RowPay.primary_payment = "-1" Then
                                    RowHeader.BayarDesc1 = "Tunai"
                                Else
                                    RowHeader.BayarDesc1 = ObjPay.GetDesc(RowPay.primary_payment)
                                End If
                                RowHeader.BayarTotal1 = RowPay.amount
                            End If
                            If IndexPayment = 1 Then
                                If RowPay.primary_payment = "-1" Then
                                    RowHeader.BayarDesc2 = "Tunai"
                                Else
                                    RowHeader.BayarDesc2 = ObjPay.GetDesc(RowPay.primary_payment)
                                End If
                                RowHeader.BayarTotal2 = RowPay.amount
                            End If
                            If IndexPayment = 2 Then
                                If RowPay.primary_payment = "-1" Then
                                    RowHeader.BayarDesc3 = "Tunai"
                                Else
                                    RowHeader.BayarDesc3 = ObjPay.GetDesc(RowPay.primary_payment)
                                End If
                                RowHeader.BayarTotal3 = RowPay.amount
                            End If
                            If IndexPayment = 3 Then
                                If RowPay.primary_payment = "-1" Then
                                    RowHeader.BayarDesc4 = "Tunai"
                                Else
                                    RowHeader.BayarDesc4 = ObjPay.GetDesc(RowPay.primary_payment)
                                End If
                                RowHeader.BayarTotal4 = RowPay.amount
                            End If
                            If IndexPayment = 4 Then
                                If RowPay.primary_payment = "-1" Then
                                    RowHeader.BayarDesc5 = "Tunai"
                                Else
                                    RowHeader.BayarDesc5 = ObjPay.GetDesc(RowPay.primary_payment)
                                End If
                                RowHeader.BayarTotal5 = RowPay.amount
                            End If
                            IndexPayment += 1
                        End If
                    Else
                        Exit For
                    End If
                Next
            End If

            'masukan data ke tabel cetakan
            For Each RowItem In TBLItem.Rows
                RowHeader = TBLHeader.Rows(0)
                Row = TBL.NewRow()
                With Row
                    'Header
                    .Tanggal = Format(RowHeader._date, "dd-MMM-yyyy")
                    .Waktu = "-"
                    .No_Transaksi = RowHeader.no_trans
                    .User = vUserCode
                    .Pelanggan = "Tunai"
                    .Grand_Total = RowHeader.grand_total
                    .Diskon_Total = RowHeader.discount
                    .Sub_Total = RowHeader.total
                    .Unit = ObjUnit.GetUnitCode(RowItem.primary_unit)
                    .Kembalian = RowHeader.change
                    'Payment
                    .Tarik_Tunai = RowHeader.CashWithdrawal
                    If RowHeader.BayarTotal1 > 0 Then
                        .PembayaranDesc1 = RowHeader.BayarDesc1
                        .PembayaranTotal1 = RowHeader.BayarTotal1
                    End If
                    If RowHeader.BayarTotal2 > 0 Then
                        .PembayaranDesc2 = RowHeader.BayarDesc2
                        .PembayaranTotal2 = RowHeader.BayarTotal2
                    End If
                    If RowHeader.BayarTotal3 > 0 Then
                        .PembayaranDesc3 = RowHeader.BayarDesc3
                        .PembayaranTotal3 = RowHeader.BayarTotal3
                    End If
                    If RowHeader.BayarTotal4 > 0 Then
                        .PembayaranDesc4 = RowHeader.BayarDesc4
                        .PembayaranTotal4 = RowHeader.BayarTotal4
                    End If
                    If RowHeader.BayarTotal5 > 0 Then
                        .PembayaranDesc5 = RowHeader.BayarDesc5
                        .PembayaranTotal5 = RowHeader.BayarTotal5
                    End If
                    'Item
                    .Item_Code = ObjProd.GetProductCode(RowItem.primary_product)
                    .Item_Desc = ObjProd.GetDescription(RowItem.primary_product)
                    .Qty = RowItem.qty
                    .Unit_Price = RowItem.unit_price
                    .Diskon_Item = RowItem.discount
                    .Total_Item = RowItem.sub_total
                End With
                TBL.Rows.Add(Row)
            Next

            Report.SetDataSource(DS)
            Dim ij As New System.Drawing.Printing.PaperSource
            ij.SourceName = "85x55"
            Report.PrintOptions.CustomPaperSource = ij
            Report.PrintOptions.PrinterName = "EPSON LX-300+ /II"
            Report.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Portrait
            Report.PrintOptions.PaperSource = CrystalDecisions.Shared.PaperSource.Tractor
            Report.ReportAppServer = "localhost"


            Dim AlamatString As String = ""
            AlamatString = "KEPENG NUSANTARA" + vbCrLf + "International Lamrim Festifal 2012" + vbCrLf + "Bandung - Indonesia"
            Report.SetParameterValue(Report.Parameter_Alamat_Pelanggan.ParameterFieldName, AlamatString)

            Row = TBL.Rows(0)
            Report.SetParameterValue(Report.Parameter_ParamStr1.ParameterFieldName, Row.PembayaranDesc1)
            Report.SetParameterValue(Report.Parameter_ParamStr2.ParameterFieldName, Row.PembayaranDesc2)
            Report.SetParameterValue(Report.Parameter_ParamStr3.ParameterFieldName, Row.PembayaranDesc3)
            Report.SetParameterValue(Report.Parameter_ParamStr4.ParameterFieldName, Row.PembayaranDesc4)
            Report.SetParameterValue(Report.Parameter_ParamStr5.ParameterFieldName, Row.PembayaranDesc5)

            Report.SetParameterValue(Report.Parameter_ParamDec1.ParameterFieldName, Row.PembayaranTotal1)
            Report.SetParameterValue(Report.Parameter_ParamDec2.ParameterFieldName, Row.PembayaranTotal2)
            Report.SetParameterValue(Report.Parameter_ParamDec3.ParameterFieldName, Row.PembayaranTotal3)
            Report.SetParameterValue(Report.Parameter_ParamDec4.ParameterFieldName, Row.PembayaranTotal4)
            Report.SetParameterValue(Report.Parameter_ParamDesc5.ParameterFieldName, Row.PembayaranTotal5)

            'setting sub total STR
            RowHeader = TBLHeader.Rows(0)
            With RowHeader
                If .discount <> 0 Then
                    'str
                    If SubTotalStr = "" Then
                        SubTotalStr = "Sub Total"
                    Else
                        SubTotalStr = SubTotalStr + "Sub Total"
                    End If
                    SubTotalStr = SubTotalStr + vbCrLf + "Diskon"
                    '---
                    If .tax_amount <> 0 Then
                        SubTotalStr = SubTotalStr + vbCrLf + "Tax"
                    End If
                    If .rounding <> 0 Then
                        SubTotalStr = SubTotalStr + vbCrLf + "Rounding"
                    End If
                    If .bank_charge <> 0 Then
                        SubTotalStr = SubTotalStr + vbCrLf + "Bank charge"
                    End If
                    '---
                    SubTotalStr = SubTotalStr + vbCrLf + "--------------------------------"
                    SubTotalStr = SubTotalStr + vbCrLf + "Total"

                    'dec
                    If SubTotalDec = "" Then
                        SubTotalDec = ObjCD.DigitGrouping(.total, 0)
                    Else
                        SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(.total, 0)
                    End If
                    SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(.discount, 0)
                    If .tax_amount <> 0 Then
                        SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(.tax_amount, 0)
                    End If
                    If .rounding <> 0 Then
                        SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(.rounding, 0)
                    End If
                    If .bank_charge <> 0 Then
                        SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(.bank_charge, 0)
                    End If
                    SubTotalDec = SubTotalDec + vbCrLf + "-------------------"
                    SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(.grand_total, 0)
                Else
                    If .discount = 0 And .tax_amount = 0 And .rounding = 0 And .bank_charge = 0 Then
                        If SubTotalStr = "" Then
                            SubTotalStr = "Total"
                        Else
                            SubTotalStr = SubTotalStr + vbCrLf + "Total"
                        End If
                        If SubTotalDec = "" Then
                            SubTotalDec = ObjCD.DigitGrouping(.grand_total, 0)
                        Else
                            SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(.grand_total, 0)
                        End If
                    Else
                        'Salin Koding
                        If SubTotalStr = "" Then
                            SubTotalStr = "Sub Total"
                        Else
                            SubTotalStr = SubTotalStr + "Sub Total"
                        End If
                        If .discount <> 0 Then
                            SubTotalStr = SubTotalStr + vbCrLf + "Diskon"
                        End If
                        If .tax_amount <> 0 Then
                            SubTotalStr = SubTotalStr + vbCrLf + "Tax"
                        End If
                        If .rounding <> 0 Then
                            SubTotalStr = SubTotalStr + vbCrLf + "Rounding"
                        End If
                        If .bank_charge <> 0 Then
                            SubTotalStr = SubTotalStr + vbCrLf + "Bank charge"
                        End If
                        SubTotalStr = SubTotalStr + vbCrLf + "--------------------------------"
                        SubTotalStr = SubTotalStr + vbCrLf + "Total"

                        'dec
                        If SubTotalDec = "" Then
                            SubTotalDec = ObjCD.DigitGrouping(.total, 0)
                        Else
                            SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(.total, 0)
                        End If
                        If .discount <> 0 Then
                            SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(.discount, 0)
                        End If
                        If .tax_amount <> 0 Then
                            SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(.tax_amount, 0)
                        End If
                        If .rounding <> 0 Then
                            SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(.rounding, 0)
                        End If
                        If .bank_charge <> 0 Then
                            SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(.bank_charge, 0)
                        End If
                        SubTotalDec = SubTotalDec + vbCrLf + "-------------------"
                        SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(.grand_total, 0)
                    End If
                End If

                'set pembayaran
                Row = TBL.Rows(0)
                If Row.PembayaranDesc2 <> "" Or Row.PembayaranDesc3 <> "" Or Row.PembayaranDesc4 <> "" Or Row.PembayaranDesc5 <> "" Then
                    SubTotalStr = SubTotalStr + vbCrLf + "Pembayaran: "
                    SubTotalDec = SubTotalDec + vbCrLf + ""
                End If
                If Row.PembayaranDesc1 <> "" Then
                    SubTotalStr = SubTotalStr + vbCrLf + Row.PembayaranDesc1
                    SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(Row.PembayaranTotal1, 0)
                End If
                If Row.PembayaranDesc2 <> "" Then
                    SubTotalStr = SubTotalStr + vbCrLf + Row.PembayaranDesc2
                    SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(Row.PembayaranTotal2, 0)
                End If
                If Row.PembayaranDesc3 <> "" Then
                    SubTotalStr = SubTotalStr + vbCrLf + Row.PembayaranDesc3
                    SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(Row.PembayaranTotal3, 0)
                End If
                If Row.PembayaranDesc4 <> "" Then
                    SubTotalStr = SubTotalStr + vbCrLf + Row.PembayaranDesc4
                    SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(Row.PembayaranTotal4, 0)
                End If
                If Row.PembayaranDesc5 <> "" Then
                    SubTotalStr = SubTotalStr + vbCrLf + Row.PembayaranDesc5
                    SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(Row.PembayaranTotal5, 0)
                End If

                'kembali
                SubTotalStr = SubTotalStr + vbCrLf + "--------------------------------"
                SubTotalDec = SubTotalDec + vbCrLf + "------------------"
                SubTotalStr = SubTotalStr + vbCrLf + "Kembalian"
                SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(.change, 0) ' - .CashWithdrawal, 0)
                If .CashWithdrawal <> 0 Then
                    SubTotalStr = SubTotalStr + vbCrLf + "Tarik Tunai"
                    SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(.CashWithdrawal, 0)
                    SubTotalStr = SubTotalStr + vbCrLf + "--------------------------------"
                    SubTotalDec = SubTotalDec + vbCrLf + "------------------"
                    SubTotalStr = SubTotalStr + vbCrLf + "Uang Diterima"
                    SubTotalDec = SubTotalDec + vbCrLf + ObjCD.DigitGrouping(.change + .CashWithdrawal, 0)
                End If
            End With
            Report.SetParameterValue(Report.Parameter_Subtotal_Str.ParameterFieldName, SubTotalStr)
            Report.SetParameterValue(Report.Parameter_SubTotalDesc.ParameterFieldName, SubTotalDec)

            vCrv.ReportSource = Report
            'If MsgBox("Cetak tanda terima?", MsgBoxStyle.YesNo, "Konfirmasi") = MsgBoxResult.Yes Then
            'Report.PrintToPrinter(1, False, 1, 15)
            'vCrv.Show()
            'vCrv.ParentForm.ShowDialog()
            'Else
            'vCrv.Show()
            'vCrv.ParentForm.ShowDialog()
            'End If
            If LangsungCetak = True Then
                Report.PrintToPrinter(1, False, 1, 1)
            Else
                If MsgBox("Cetak tanda terima?", MsgBoxStyle.YesNo, "Konfirmasi") = MsgBoxResult.Yes Then
                    vCrv.Show()
                    vCrv.ParentForm.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTAHeader.Dispose()
            DTAItem.Dispose()
            DTAPay.Dispose()
            TBLHeader.Dispose()
            TBLItem.Dispose()
            TBLPay.Dispose()
            ObjProd = Nothing
            ObjUnit = Nothing
            ObjPay = Nothing
            ObjCD = Nothing
        End Try
    End Sub

    Public Sub BackUp_CetakPOSUkuranStngahLetter(ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vUserCode As String, ByVal vPrimary As Double, Optional ByVal LangsungCetak As Boolean = False)
        Dim Report As New BackUp_CrpCetakPOSSetengahLetter
        Dim DS As New DseTrans_POS_Cetak
        Dim TBL As DseTrans_POS_Cetak.Tabel_CetakDataTable = DS.Tabel_Cetak
        Dim Row As DseTrans_POS_Cetak.Tabel_CetakRow
        Dim ObjProd As New clsMasterProduct
        Dim ObjUnit As New ClsMasterUnit
        Dim ObjPay As New ClsMast_PaymentOption
        Dim DTAHeader As New DseTrans_POS_CetakTableAdapters.tbltrans_pos_cabangTableAdapter
        Dim DTAItem As New DseTrans_POS_CetakTableAdapters.tbltrans_pos_cabang_itemTableAdapter
        Dim DTAPay As New DseTrans_POS_CetakTableAdapters.tbltrans_pos_cabang_item_paymentTableAdapter
        Dim TBLHeader As New DseTrans_POS_Cetak.tbltrans_pos_cabangDataTable
        Dim TBLItem As New DseTrans_POS_Cetak.tbltrans_pos_cabang_itemDataTable
        Dim TBLPay As New DseTrans_POS_Cetak.tbltrans_pos_cabang_item_paymentDataTable
        Dim RowHeader As DseTrans_POS_Cetak.tbltrans_pos_cabangRow
        Dim RowItem As DseTrans_POS_Cetak.tbltrans_pos_cabang_itemRow
        Dim RowPay As DseTrans_POS_Cetak.tbltrans_pos_cabang_item_paymentRow
        Dim IndexPayment As Integer = 0
        Dim JumlahTarikTunai As Decimal = 0
        Try
            DTAHeader.Connection.ConnectionString = Me.GetConstring
            DTAItem.Connection.ConnectionString = Me.GetConstring
            DTAPay.Connection.ConnectionString = Me.GetConstring

            DTAHeader.FillBy_Primary(TBLHeader, vPrimary)
            DTAItem.FillBy_Primary(TBLItem, vPrimary)
            DTAPay.FillBy_Primary(TBLPay, vPrimary)

            'Hitung dulu payment-nya 
            RowHeader = TBLHeader.Rows(0)
            If RowHeader.is_single_payment = 0 Then
                RowPay = TBLPay.Rows(0)
                If RowPay.primary_payment = "-1" Then
                    RowHeader.BayarDesc1 = "Tunai"
                    RowHeader.BayarTotal1 = RowPay.amount
                Else
                    RowHeader.BayarDesc1 = ObjPay.GetDesc(RowPay.primary_payment)
                    RowHeader.BayarTotal1 = RowPay.amount
                End If
            Else
                For Each RowPay In TBLPay.Rows
                    If IndexPayment <= 4 Then
                        If ObjPay.GetTypePayment(RowPay.primary_payment) = 2 Then
                            RowHeader.CashWithdrawal += RowPay.amount
                        Else
                            If IndexPayment = 0 Then
                                If RowPay.primary_payment = "-1" Then
                                    RowHeader.BayarDesc1 = "Tunai"
                                Else
                                    RowHeader.BayarDesc1 = ObjPay.GetDesc(RowPay.primary_payment)
                                End If
                                RowHeader.BayarTotal1 = RowPay.amount
                            End If
                            If IndexPayment = 1 Then
                                If RowPay.primary_payment = "-1" Then
                                    RowHeader.BayarDesc2 = "Tunai"
                                Else
                                    RowHeader.BayarDesc2 = ObjPay.GetDesc(RowPay.primary_payment)
                                End If
                                RowHeader.BayarTotal2 = RowPay.amount
                            End If
                            If IndexPayment = 2 Then
                                If RowPay.primary_payment = "-1" Then
                                    RowHeader.BayarDesc3 = "Tunai"
                                Else
                                    RowHeader.BayarDesc3 = ObjPay.GetDesc(RowPay.primary_payment)
                                End If
                                RowHeader.BayarTotal3 = RowPay.amount
                            End If
                            If IndexPayment = 3 Then
                                If RowPay.primary_payment = "-1" Then
                                    RowHeader.BayarDesc4 = "Tunai"
                                Else
                                    RowHeader.BayarDesc4 = ObjPay.GetDesc(RowPay.primary_payment)
                                End If
                                RowHeader.BayarTotal4 = RowPay.amount
                            End If
                            If IndexPayment = 4 Then
                                If RowPay.primary_payment = "-1" Then
                                    RowHeader.BayarDesc5 = "Tunai"
                                Else
                                    RowHeader.BayarDesc5 = ObjPay.GetDesc(RowPay.primary_payment)
                                End If
                                RowHeader.BayarTotal5 = RowPay.amount
                            End If
                            IndexPayment += 1
                        End If
                    Else
                        Exit For
                    End If
                Next
            End If

            'masukan data ke tabel cetakan
            For Each RowItem In TBLItem.Rows
                RowHeader = TBLHeader.Rows(0)
                Row = TBL.NewRow()
                With Row
                    'Header
                    .Tanggal = Format(RowHeader._date, "dd-MMM-yyyy")
                    .Waktu = "-"
                    .No_Transaksi = RowHeader.no_trans
                    .User = vUserCode
                    .Pelanggan = "Tunai"
                    .Grand_Total = RowHeader.grand_total
                    .Diskon_Total = RowHeader.discount
                    .Sub_Total = RowHeader.total
                    .Unit = ObjUnit.GetUnitCode(RowItem.primary_unit)
                    .Kembalian = RowHeader.change
                    'Payment
                    .Tarik_Tunai = RowHeader.CashWithdrawal
                    If RowHeader.BayarTotal1 > 0 Then
                        .PembayaranDesc1 = RowHeader.BayarDesc1
                        .PembayaranTotal1 = RowHeader.BayarTotal1
                    End If
                    If RowHeader.BayarTotal2 > 0 Then
                        .PembayaranDesc2 = RowHeader.BayarDesc2
                        .PembayaranTotal2 = RowHeader.BayarTotal2
                    End If
                    If RowHeader.BayarTotal3 > 0 Then
                        .PembayaranDesc3 = RowHeader.BayarDesc3
                        .PembayaranTotal3 = RowHeader.BayarTotal3
                    End If
                    If RowHeader.BayarTotal4 > 0 Then
                        .PembayaranDesc4 = RowHeader.BayarDesc4
                        .PembayaranTotal4 = RowHeader.BayarTotal4
                    End If
                    If RowHeader.BayarTotal5 > 0 Then
                        .PembayaranDesc5 = RowHeader.BayarDesc5
                        .PembayaranTotal5 = RowHeader.BayarTotal5
                    End If
                    'Item
                    .Item_Code = ObjProd.GetProductCode(RowItem.primary_product)
                    .Item_Desc = ObjProd.GetDescription(RowItem.primary_product)
                    .Qty = RowItem.qty
                    .Unit_Price = RowItem.unit_price
                    .Diskon_Item = RowItem.discount
                    .Total_Item = RowItem.sub_total
                End With
                TBL.Rows.Add(Row)
            Next
            Dim ij As New System.Drawing.Printing.PaperSource
            Report.SetDataSource(DS)
            Report.PrintOptions.PrinterName = "EPSON LX-300+ /II"
            ij.SourceName = "85x55"
            Report.PrintOptions.CustomPaperSource = ij
            Report.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Portrait
            Report.PrintOptions.PaperSource = CrystalDecisions.Shared.PaperSource.Tractor
            Report.ReportAppServer = "localhost"

            Dim AlamatString As String = ""
            AlamatString = "PT. ABADI CAHAYA PURNAMA" + vbCrLf + "Jl. Abdurahman Saleh No. 9" + vbCrLf + "Ruko City Square B 12" + vbCrLf + "Bandung - 40123"
            Report.SetParameterValue(Report.Parameter_Alamat_Pelanggan.ParameterFieldName, AlamatString)

            Row = TBL.Rows(0)
            Report.SetParameterValue(Report.Parameter_ParamStr1.ParameterFieldName, Row.PembayaranDesc1)
            Report.SetParameterValue(Report.Parameter_ParamStr2.ParameterFieldName, Row.PembayaranDesc2)
            Report.SetParameterValue(Report.Parameter_ParamStr3.ParameterFieldName, Row.PembayaranDesc3)
            Report.SetParameterValue(Report.Parameter_ParamStr4.ParameterFieldName, Row.PembayaranDesc4)
            Report.SetParameterValue(Report.Parameter_ParamStr5.ParameterFieldName, Row.PembayaranDesc5)

            Report.SetParameterValue(Report.Parameter_ParamDec1.ParameterFieldName, Row.PembayaranTotal1)
            Report.SetParameterValue(Report.Parameter_ParamDec2.ParameterFieldName, Row.PembayaranTotal2)
            Report.SetParameterValue(Report.Parameter_ParamDec3.ParameterFieldName, Row.PembayaranTotal3)
            Report.SetParameterValue(Report.Parameter_ParamDec4.ParameterFieldName, Row.PembayaranTotal4)
            Report.SetParameterValue(Report.Parameter_ParamDesc5.ParameterFieldName, Row.PembayaranTotal5)

            vCrv.ReportSource = Report
            If LangsungCetak = True Then
                Report.PrintToPrinter(1, False, 0, 0)
            Else
                vCrv.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTAHeader.Dispose()
            DTAItem.Dispose()
            DTAPay.Dispose()
            TBLHeader.Dispose()
            TBLItem.Dispose()
            TBLPay.Dispose()
            ObjProd = Nothing
            ObjUnit = Nothing
            ObjPay = Nothing
        End Try
    End Sub

    Public Function GetJmlByPrimaryMember(ByVal vPrimaryMember As Double) As Decimal
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTa.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryMember(vPrimaryMember).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimaryMember(vPrimaryMember).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByPrimaryDiskon(ByVal vPrimaryDiskon As Double) As Decimal
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryDiscount(vPrimaryDiskon).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimaryDiscount(vPrimaryDiskon).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByPrimarySession(ByVal vPrimarySession As Double) As Decimal
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimarySession(vPrimarySession).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimarySession(vPrimarySession).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByPrimaryPayment(ByVal vPrimaryPayment As Double) As Decimal
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabang_item_paymentTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryPayment(vPrimaryPayment).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimaryPayment(vPrimaryPayment).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimarySession(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimarySession(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetPrimarySession(vPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryMember(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimaryMember(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryMember(vPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CekApakahSudahPaid(ByVal vPrimary As Double) As Decimal
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If vPrimary <> -1 Then
                If DTA.ScalarQuery_GetTotalPayment(vPrimary).Value > 0 Then
                    Return True
                Else
                    Return False
                End If
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetChange(ByVal vPrimary As Double) As Decimal
        Dim DTA As New DseTrans_POSTableAdapters.tbltrans_pos_cabangTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJumlahChange(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetJumlahChange(vPrimary).Value
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
