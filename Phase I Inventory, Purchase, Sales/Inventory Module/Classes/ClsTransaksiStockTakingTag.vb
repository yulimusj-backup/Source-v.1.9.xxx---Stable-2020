Option Explicit On
Option Strict Off

Public Class ClsTransaksiStockTakingTag
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mPrimaryStockTaking As Double = -1
    Dim mTagNo As String = ""
    Dim mTeamIdentifier As String = ""
    Dim mTanggalTransaksi As Date = Now
    Dim mPrimaryLocation As Double = -1
    Dim mRemarks As String = ""
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Property TeamIdentifier() As String
        Get
            Return Me.mTeamIdentifier
        End Get
        Set(ByVal value As String)
            mTeamIdentifier = value
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

    Public Property PrimaryLocation() As Double
        Get
            Return Me.mPrimaryLocation
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryLocation = value
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
            Return Me.mTagNo
        End Get
        Set(ByVal value As String)
            Me.mTagNo = value
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

    Public Property PrimaryStockTaking() As Double
        Get
            Return Me.mPrimaryStockTaking
        End Get
        Set(ByVal value As Double)
            mPrimaryStockTaking = value
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

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String, _
    ByVal vPrimaryStockTaking As Double) As Boolean
        Dim DTA As New DseTransaksiStockTagTableAdapters.tbltrans_stock_taking_tagTableAdapter

        Try
            If DTA.ScalarQuery_GetJml(vPrimaryStockTaking, vNoTransaksi).HasValue = True Then
                If DTA.ScalarQuery_GetJml(vPrimaryStockTaking, vNoTransaksi).Value > 0 Then
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
        Dim DTA As New DseTransaksiStockTagTableAdapters.tbltrans_stock_taking_tagTableAdapter
        Dim TBL As New DseTransaksiStockTag.tbltrans_stock_taking_tagDataTable
        Dim row As DseTransaksiStockTag.tbltrans_stock_taking_tagRow
        Dim ObjItem As New clsTransaksiStockTakingTagItem

        Try
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mTagNo = .Tag_No
                    Me.mPrimary = .Primary
                    Me.mPrimaryStockTaking = .Primary_Stock_Taking
                    Me.mPrimaryLocation = .Primary_Storage_Area
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = ._Date
                    Me.mTeamIdentifier = .Team_Identifier

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

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryStockTaking As Double) As Boolean
        Dim Obj As New clsTransaksiStockTaking

        Try
            'tidak boleh diedit jika transaksi stock taking yang dirujuk sudah diproses, atau tidak ada
            'transaksi stock taking yang dirujuk
            If IsNothing(Obj.GetStatusStockTaking(vPrimaryStockTaking)) = False Then
                If Obj.GetStatusStockTaking(vPrimaryStockTaking) = clsTransaksiStockTaking.EnumStatusStockTaking.Active Then
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
            Obj = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryStockTaking As Double) As Boolean
        Dim Obj As New clsTransaksiStockTaking

        Try
            'tidak boleh dihapus jika transaksi stock taking yang dirujuk sudah diproses, atau tidak ada
            'transaksi stock taking yang dirujuk
            If IsNothing(Obj.GetStatusStockTaking(vPrimaryStockTaking)) = False Then
                If Obj.GetStatusStockTaking(vPrimaryStockTaking) = clsTransaksiStockTaking.EnumStatusStockTaking.Active Then
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
            Obj = Nothing
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean, Optional ByVal vCekDataBenar As Boolean = True) As Boolean
        Try
            If vCekDataBenar = False Then
                Return True
            End If
            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.

            '1.b No transaksi kosong

            '2. Jika semua transaksi inventory sedang dikunci

            '3. Jika timestamp yang dipegang user sudah tidak valid lagi

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, _
    ByVal vUserCode As String, Optional ByVal vCekDataBenar As Boolean = True, _
    Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiStockTagTableAdapters.tbltrans_stock_taking_tagTableAdapter
        Dim ObjItem As New clsTransaksiStockTakingTagItem

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim Timestamp_ As Date = Now

        Try
            If Me.ApakahDataBenar(True, vCekDataBenar) = True Then
                If vFromCutOff = True Then
                    CreatedBy_ = Me.mCreatedBy
                    CreatedDate_ = Me.mCreatedDate
                    ModifiedBy_ = Me.mModifiedBy
                    ModifiedDate_ = Me.mModifiedDate
                    Timestamp_ = Me.mTimestamp
                End If

                'insert query
                If DTA.Insert(Me.mPrimaryStockTaking, Me.mTagNo, Me.mTeamIdentifier, _
                Me.mTanggalTransaksi, Me.mPrimaryLocation, Me.mRemarks, CreatedBy_, _
                CreatedDate_, ModifiedBy_, ModifiedDate_, Timestamp_) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mTagNo, Me.mPrimaryStockTaking).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    'simpan data item baru
                    ObjItem.Simpan(vdgvItem, Me.mTanggalTransaksi, Me.mPrimary, Me.mTagNo)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiStockTagTableAdapters.tbltrans_stock_taking_tagTableAdapter
        Dim ObjItem As New clsTransaksiStockTakingTagItem

        Try
            If Me.ApakahDataBenar(False) = True Then
                If DTA.Update(Me.mPrimaryStockTaking, Me.mTagNo, Me.mTeamIdentifier, Me.mTanggalTransaksi, _
                Me.mPrimaryLocation, Me.mRemarks, vUserCode, Now, Now, Me.mPrimary, _
                Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItem, Me.mTanggalTransaksi, Me.mPrimary, _
                    Me.mTagNo)
                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim Obj As New clsTransaksiStockTakingTagItem
        Dim DTA As New DseTransaksiStockTagTableAdapters.tbltrans_stock_taking_tagTableAdapter

        Try
            If Me.ApakahBolehHapus(vPrimaryTransaksi) = True Then
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

    Public Function GetJmlByPrimaryStockTaking(ByVal vPrimaryStockTaking As Double) As Double
        Dim DTa As New DseTransaksiStockTagTableAdapters.tbltrans_stock_taking_tagTableAdapter

        Try
            If DTa.ScalarQuery_GetJmlByPrimaryStockTaking(vPrimaryStockTaking).HasValue = True Then
                Return DTa.ScalarQuery_GetJmlByPrimaryStockTaking(vPrimaryStockTaking).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTa.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
    ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiStockTagTableAdapters.tbltrans_stock_taking_tagTableAdapter

        Try
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                If DTA.ScalarQuery_GetJmlByTagNoLIKE(Criteria + "%").HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByTagNoLIKE(Criteria + "%")
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                If DTA.ScalarQuery_GetJmlByTanggalSampaiDari(vTanggalSampai, vTanggalDari).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByTanggalSampaiDari(vTanggalSampai, vTanggalDari).Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                If DTA.ScalarQuery_GetJmlByKodeProdukLIKE(Criteria + "%").HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByKodeProdukLIKE(Criteria + "%")
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                If DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%").HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%")
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByTakingTeamNo Then
                If DTA.ScalarQuery_GetJmlByTeamNoLIKE(Criteria + "%").HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByTeamNoLIKE(Criteria + "%")
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoWorkOrder Then
                'msgbox("belum")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiStockTag.trans_stocktaking_tag_viewDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiStockTagTableAdapters.trans_stocktaking_tag_viewTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_TagNoLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TanggalSampaiDari(TBL, vTanggalSampai, vTanggalDari)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescriptionLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTakingTeamNo Then
                QueryResult = DTA.FillBy_TeamNoLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Transaction No"
                DGV.Columns(4).DefaultCellStyle.Format = "dd MMM yyyy"
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

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            Me.mCreatedBy = ""
            Me.mCreatedDate = Now
            Me.mModifiedBy = ""
            Me.mModifiedDate = Now
            Me.mTagNo = ""
            Me.mPrimary = -1
            Me.mPrimaryStockTaking = -1
            Me.mRemarks = ""
            Me.mTanggalTransaksi = Now
            Me.mTimestamp = Now
            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
