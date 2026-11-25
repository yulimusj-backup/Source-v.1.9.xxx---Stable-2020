Option Explicit On
Option Strict Off

Public Class clsTransaksiStockTaking
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mNoTransaksi As String = ""
    Dim mStartDate As Date = Now
    Dim mDesc As String = ""
    Dim mTotalTeam As Integer = 0
    Dim mStatus As EnumStatusStockTaking = EnumStatusStockTaking.Active
    ''' <summary>
    ''' tanggal dilakukan proses perubahan data
    ''' </summary>
    ''' <remarks></remarks>
    Dim mProcessDate As Date = Now

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Property ProcessDate() As Date
        Get
            Return mProcessDate
        End Get
        Set(ByVal value As Date)
            mProcessDate = value
        End Set
    End Property

    Public Property Status() As EnumStatusStockTaking
        Get
            Return mStatus
        End Get
        Set(ByVal value As EnumStatusStockTaking)
            mStatus = value
        End Set
    End Property

    Public Property TotalTeam() As Integer
        Get
            Return mTotalTeam
        End Get
        Set(ByVal value As Integer)
            mTotalTeam = value
        End Set
    End Property
    Public Property Desc() As String
        Get
            Return mDesc
        End Get
        Set(ByVal value As String)
            mDesc = value
        End Set
    End Property
    Public Property NoTransaksi() As String
        Get
            Return mNoTransaksi
        End Get
        Set(ByVal value As String)
            mNoTransaksi = value
        End Set
    End Property
    Public Property StartDate() As Date
        Get
            Return mStartDate
        End Get
        Set(ByVal value As Date)
            mStartDate = value
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

    Public Enum EnumStatusStockTaking
        Active
        Processed
    End Enum

    Public Function SetStatus(ByVal vPrimaryStocktaking As Double, ByVal Status As EnumStatusStockTaking) As Boolean
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Try
            DTA.UpdateQuery_SetStatus(Status, vPrimaryStocktaking)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetStatus(ByVal vPrimaryStockTaking As Double) As EnumStatusStockTaking
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Try
            If DTA.ScalarQuery_GetStatus(vPrimaryStockTaking).HasValue Then
                Return DTA.ScalarQuery_GetStatus(vPrimaryStockTaking).Value
            Else
                Return Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiTBL()

    End Sub

    Public Function IsiDGV(ByVal vNoTransaksi As String, ByVal vDescription As String, _
    ByVal DGV As DataGridView, ByVal TBL As dseTransaksiStockTaking.tbltrans_stock_taking2DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, Optional ByVal vStartDate As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking2TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_StartDate(TBL, vStartDate)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_DescLike(TBL, vDescription + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLike(TBL, vNoTransaksi + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                Dim ij As New DataGridViewCellStyle
                ij.Format = "dd MMM yyyy"
                DGV.Columns(2).DefaultCellStyle = ij
                DGV.Columns(1).DefaultCellStyle = ij
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
    ''' <summary>
    ''' mengisikan parameter LS dengan nomor team taking yang merujuk ke primary stock taking tertentu.
    ''' </summary>
    ''' <param name="vPrimaryStockTaking"></param>
    ''' <param name="LS"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTeamNoList(ByVal vPrimaryStockTaking As Double, _
    ByVal LS As List(Of String)) As Boolean
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking_teamTableAdapter
        Dim TBL As dseTransaksiStockTaking.tbltrans_stock_taking_teamDataTable
        Dim row As dseTransaksiStockTaking.tbltrans_stock_taking_teamRow
        Try
            TBL = DTA.GetDataBy_PrimaryStockTaking(vPrimaryStockTaking)
            If TBL.Rows.Count > 0 Then
                LS.Clear()
                For Each row In TBL.Rows
                    If LS.Contains(row.Team_No) = False Then
                        LS.Add(row.Team_No)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlActiveStockTaking() As Double
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Dim i As EnumStatusStockTaking = EnumStatusStockTaking.Active
        Try
            If DTA.ScalarQuery_GetJmlByStatus(i).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByStatus(i).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlActiveStockTaking(ByVal vStartDate As Date, ByVal vEndDate As Date) As Double
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Dim i As EnumStatusStockTaking = EnumStatusStockTaking.Active

        Try
            If DTA.ScalarQuery_GetJml2(i, vEndDate, vStartDate).HasValue = True Then
                Return DTA.ScalarQuery_GetJml2(i, vEndDate, vStartDate).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetStartDateTransaksiYangMasihAktif() As Date
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Try
            If DTA.ScalarQuery_getStartDateStockTaking().HasValue = True Then
                Return DTA.ScalarQuery_getStartDateStockTaking().Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'mengembalikan jumlah stock taking, berdasarkan tanggal start, yang primary tidak sama dengan primary
    'yang diberikan
    Public Function GetJml(ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vPrimaryStockTaking As Double) As Double
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Dim Dummy As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJml3(vEndDate, vStartDate, vPrimaryStockTaking)
            If Dummy.HasValue = True Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vStartDate As Date, ByVal vEndDate As Date) As Double
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Dim Dummy As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJml(vEndDate, vStartDate)
            If Dummy.HasValue = True Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vNoTransaksi As String, ByVal vDescription As String, _
        ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vStartDate As Date = (#1/1/1950#)) As Long
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Try
            If SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                If DTA.ScalarQuery_getJmlDatabyStartDate(vStartDate).HasValue = True Then
                    Return DTA.ScalarQuery_getJmlDatabyStartDate(vStartDate).Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                If DTA.ScalarQuery_getJmlDataByDescriptionLike(vDescription + "%").HasValue = True Then
                    Return DTA.ScalarQuery_getJmlDataByDescriptionLike(vDescription + "%").Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                If DTA.ScalarQuery_getJmlbyNoTransaksiLike(vNoTransaksi + "%").HasValue = True Then
                    Return DTA.ScalarQuery_getJmlbyNoTransaksiLike(vNoTransaksi + "%").Value
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

    Public Function GetPrimaryMax() As Double
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Try
            If DTA.ScalarQuery_GetMaxPrimary.HasValue Then
                Return DTA.ScalarQuery_GetMaxPrimary.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTeamNo(ByVal vTeamIdentifier As String) As String
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking_teamTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetTeamNo(vTeamIdentifier)) = False Then
                Return DTA.ScalarQuery_GetTeamNo(vTeamIdentifier)
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

    Public Function GetTeamIdentifier(ByVal TeamNo As String, ByVal vPrimaryStockTaking As Double) As String
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking_teamTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetIdentifier(TeamNo, vPrimaryStockTaking)) = False Then
                Return DTA.ScalarQuery_GetIdentifier(TeamNo, vPrimaryStockTaking)
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

    Public Function IsiCboTeamNo(ByVal vPrimaryTransaksi As Double, _
    ByVal TBL As dseTransaksiStockTaking.tbltrans_stock_taking_teamDataTable, ByVal Cbo As ComboBox) _
    As Boolean
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking_teamTableAdapter

        Try
            If DTA.FillBy_PrimaryStockTaking(TBL, vPrimaryTransaksi) > 0 Then
                Cbo.DataSource = TBL.DefaultView
                Cbo.DisplayMember = "Team No"
                Cbo.SelectedIndex = 0
            Else
                Cbo.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Try
            If DTA.ScalarQuery_getPrimaryByNoTransaksi(vNoTransaksi).HasValue Then
                Return DTA.ScalarQuery_getPrimaryByNoTransaksi(vNoTransaksi).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetStartDate(ByVal vPrimary As Double) As Date
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter

        Try
            If DTA.ScalarQuery_GetStartDate(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetStartDate(vPrimary).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahTermasukTanggal(ByVal TanggalInventoryCheckTag As Date, _
    ByVal vPrimaryStockTaking As Double, Optional ByRef StartDate As Date = #1/1/1950#) As Boolean
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Dim TBL As New dseTransaksiStockTaking.tbltrans_stock_takingDataTable
        Dim row As dseTransaksiStockTaking.tbltrans_stock_takingRow

        Try
            If DTA.FillBy_Primary(TBL, vPrimaryStockTaking) > 0 Then
                row = TBL.Rows(0)
                StartDate = row.Start_Date

                If TanggalInventoryCheckTag = row.Start_Date Then
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
            TBL.Dispose()
        End Try
    End Function

    Public Function CekApakahAdaYangMasihAktif(Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_getDescriptionByStatus(clsTransaksiStockTaking.EnumStatusStockTaking.Active)) = True Then
                Return False
            Else
                If TampilkanPesan = True Then
                    MsgBox("Transaction " + _
                    DTA.ScalarQuery_GetNoTransaksiByStatus(clsTransaksiStockTaking.EnumStatusStockTaking.Active) + _
                    " is still active. There must be only one active stock taking transaction at a time.", _
                    MsgBoxStyle.Information, "Information")
                End If

                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengambil semua Identifier dari tbltransaksi_stock_taking_team
    ''' </summary>
    ''' <param name="vPrimaryStockTaking"></param>
    ''' <param name="LS"></param>
    ''' <remarks></remarks>
    Public Sub GetListIdentifier(ByVal vPrimaryStockTaking As String, ByVal LS As List(Of String))
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking_teamTableAdapter
        Dim TBL As dseTransaksiStockTaking.tbltrans_stock_taking_teamDataTable
        Dim row As dseTransaksiStockTaking.tbltrans_stock_taking_teamRow
        Try
            TBL = DTA.GetDataBy_PrimaryStockTaking(vPrimaryStockTaking)
            If TBL.Rows.Count > 0 Then
                LS.Clear()
                For Each row In TBL.Rows
                    If LS.Contains(row.Identifier) = False Then
                        LS.Add(row.Identifier)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function ApakahBolehHapus(ByVal vPrimaryStockTaking As Double) As Boolean
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Dim Obj As New ClsTransaksiStockTakingTag
        Dim Status As EnumStatusStockTaking

        Try
            'tidak boleh hapus jika:
            '1. Transaksi sudah diproses
            If DTA.ScalarQuery_GetStatus(vPrimaryStockTaking).HasValue = True Then
                Status = DTA.ScalarQuery_GetStatus(vPrimaryStockTaking).Value
            Else
                Status = EnumStatusStockTaking.Processed
            End If
            If Status = EnumStatusStockTaking.Processed Then
                MsgBox("This transaction has been processed, you can not delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            '2. Sudah ada stock tag yang merujuk pada transaksi stock taking ybs
            If Obj.GetJmlByPrimaryStockTaking(vPrimaryStockTaking) > 0 Then
                MsgBox("This transaction referred by one or more inventory check tag, you can not delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'kalo sampe sini berarti sudah boleh hapus
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Dim Dummy As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi)
            If Dummy.HasValue = True Then
                If Dummy.Value > 0 Then
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

    Public Function ApakahBolehEdit(ByVal vPrimaryStockTaking As Double) As Boolean
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Dim Status As EnumStatusStockTaking

        Try
            'tidak boleh hapus jika:
            '1. Transaksi sudah diproses
            If DTA.ScalarQuery_GetStatus(vPrimaryStockTaking).HasValue = True Then
                Status = DTA.ScalarQuery_GetStatus(vPrimaryStockTaking).Value
            Else
                Status = EnumStatusStockTaking.Processed
            End If
            If Status = EnumStatusStockTaking.Processed Then
                MsgBox("This transaction has been processed, you can not edit this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'kalo sampe sini berarti sudah boleh hapus
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetStatusStockTaking(ByVal vprimary As String) As EnumStatusStockTaking
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Try
            If DTA.ScalarQuery_GetStatus(vprimary).HasValue = True Then
                Return DTA.ScalarQuery_GetStatus(vprimary).Value
            Else
                Return Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub ResetObject()
        Try
            Me.mCreatedBy = ""
            Me.mNoTransaksi = ""
            Me.mCreatedDate = Now
            Me.mModifiedBy = ""
            Me.mModifiedDate = Now
            Me.mPrimary = -1
            Me.mProcessDate = Now
            Me.mStartDate = Now
            Me.mDesc = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function Cetak_ExpiryDataCheck(ByVal vReport As crpLaporanStockTaking_ED_DataCheck, ByVal vCrv As _
    CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal VprimaryTransaksi As Double) As Boolean
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal TabCont As TabControl) As Boolean
        Dim DTATeam As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking_teamTableAdapter
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        '     Dim DTAStaff As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking_staffTableAdapter
        Try
            'cari data header
            Dim TBL As dseTransaksiStockTaking.tbltrans_stock_takingDataTable
            Dim row As dseTransaksiStockTaking.tbltrans_stock_takingRow
            TBL = DTA.GetDataBy_Primary(vPrimary)
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    Me.Primary = row.Primary
                    Me.NoTransaksi = row.No_Transaksi
                    Me.CreatedBy = row.Created_By
                    Me.CreatedDate = row.Created_Date
                    Me.Desc = row.Description
                    Me.ModifiedBy = row.Modified_By
                    Me.ModifiedDate = row.Modified_Date
                    Me.StartDate = row.Start_Date
                    Me.Status = row.Status
                    Me.TimeStamp = row.Timestamp
                    Me.TotalTeam = row.Total_Team
                Next
                'masukkan data item dan modifikasi tab control jika diperlukan
                Do While TabCont.TabPages.Count <> 1
                    TabCont.TabPages.RemoveAt(TabCont.TabPages.Count - 1)
                Loop
                Dim TBLTeam As dseTransaksiStockTaking.tbltrans_stock_taking_teamDataTable
                Dim RowTeam As dseTransaksiStockTaking.tbltrans_stock_taking_teamRow
                TBLTeam = DTATeam.GetDataBy_PrimaryStockTaking(vPrimary)
                If TBLTeam.Rows.Count > 0 Then
                    For Each RowTeam In TBLTeam.Rows
                        Me.TambahTabPageDanAccesoriesnya(TabCont, RowTeam)
                    Next
                End If
                'nilai balikan
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTATeam.Dispose()
            DTA.Dispose()
            '        DTAStaff.Dispose()
        End Try
    End Function
    Private Sub TambahTabPageDanAccesoriesnya(ByVal TabCont As TabControl, ByVal RowTeam As dseTransaksiStockTaking.tbltrans_stock_taking_teamRow)
        Try
            Dim i As New TabPage
            'lbTeamNo
            Dim lbTeamNo As New Label
            lbTeamNo.AutoSize = True
            lbTeamNo.Location = New System.Drawing.Point(8, 9)
            lbTeamNo.Name = "lbTeamNo" + CStr(TabCont.TabPages.Count)
            lbTeamNo.Size = New System.Drawing.Size(51, 13)
            lbTeamNo.TabIndex = 47
            lbTeamNo.Text = "Team No"
            'Tambah Text Box
            Dim txtTeamNumber As New clsTxtCustom
            txtTeamNumber.Location = New System.Drawing.Point(64, 6)
            txtTeamNumber.Name = "txtTeamNumber" + CStr(TabCont.TabPages.Count)
            txtTeamNumber.Size = New System.Drawing.Size(88, 20)
            txtTeamNumber.TabIndex = 48
            txtTeamNumber.Text = RowTeam.Team_No
            'lbJobDesc
            Dim lbJobDesc As New Label
            lbJobDesc.AutoSize = True
            lbJobDesc.Location = New System.Drawing.Point(8, 29)
            lbJobDesc.Name = "lbJobDesc" + CStr(TabCont.TabPages.Count)
            lbJobDesc.Size = New System.Drawing.Size(80, 13)
            lbJobDesc.TabIndex = 49
            lbJobDesc.Text = "Job Description"
            'txtTeamJobDesc
            Dim txtTeamJobDesc As New clsTxtCustom
            txtTeamJobDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                        Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
            txtTeamJobDesc.Location = New System.Drawing.Point(11, 45)
            txtTeamJobDesc.Multiline = True
            txtTeamJobDesc.Name = "txtTeamJobDesc" + CStr(TabCont.TabPages.Count)
            txtTeamJobDesc.Size = New System.Drawing.Size(141, 50)
            txtTeamJobDesc.TabIndex = 50
            txtTeamJobDesc.Text = RowTeam.Job_Description
            'Tambah DGV
            Dim DgvTeam As New ClsDGVCustom
            Me.TambahDGV(DgvTeam, TabCont, RowTeam)
            'gboTeamStaff
            Dim gboTeamStaff As New GroupBox
            gboTeamStaff.Controls.Add(DgvTeam)
            gboTeamStaff.Location = New System.Drawing.Point(158, 6)
            gboTeamStaff.Name = "gboTeamStaff" + CStr(TabCont.TabPages.Count)
            gboTeamStaff.Size = New System.Drawing.Size(356, 202)
            gboTeamStaff.TabIndex = 51
            gboTeamStaff.TabStop = False
            gboTeamStaff.Text = "Team Staff"


            'Masukan Item pada Tab Page
            i.Controls.Add(gboTeamStaff)
            i.Controls.Add(txtTeamJobDesc)
            i.Controls.Add(lbJobDesc)
            i.Controls.Add(txtTeamNumber)
            i.Controls.Add(lbTeamNo)
            i.Location = New System.Drawing.Point(4, 22)
            i.Name = "TabPage" + CStr(TabCont.TabPages.Count)
            i.Padding = New System.Windows.Forms.Padding(3)
            i.Size = New System.Drawing.Size(520, 214)
            i.TabIndex = 1
            i.Text = "Team " + CStr(TabCont.TabPages.Count)
            i.UseVisualStyleBackColor = True

            TabCont.TabPages.Add(i)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahDGV(ByVal DgvTeam As ClsDGVCustom, ByVal TabCont As TabControl, ByVal RowTeam As dseTransaksiStockTaking.tbltrans_stock_taking_teamRow)
        Dim DTAStaff As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking_staffTableAdapter
        Try
            Dim Column1 As New System.Windows.Forms.DataGridViewTextBoxColumn
            Dim Column2 As New System.Windows.Forms.DataGridViewTextBoxColumn
            Dim Column3 As New System.Windows.Forms.DataGridViewTextBoxColumn
            '
            'Column1
            '
            Column1.HeaderText = "No"
            Column1.Name = "ColNo" + CStr(TabCont.TabPages.Count - 1)
            Column1.Width = 46
            '
            'Column2
            '
            Column2.HeaderText = "Name"
            Column2.Name = "ColName" + CStr(TabCont.TabPages.Count - 1)
            Column2.Width = 60
            '
            'Column3
            '
            Column3.HeaderText = "Remark"
            Column3.Name = "ColRemark" + CStr(TabCont.TabPages.Count - 1)
            Column3.Width = 69
            '
            DgvTeam.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
            DgvTeam.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            DgvTeam.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Column1, Column2, Column3})
            DgvTeam.Dock = System.Windows.Forms.DockStyle.Fill
            DgvTeam.Location = New System.Drawing.Point(3, 16)
            DgvTeam.Name = "DataGridView" + CStr(TabCont.TabPages.Count)
            DgvTeam.Size = New System.Drawing.Size(350, 183)
            DgvTeam.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
            DgvTeam.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            DgvTeam.MultiSelect = False
            DgvTeam.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DgvTeam.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            DgvTeam.AllowUserToOrderColumns = False
            DgvTeam.AllowUserToResizeColumns = False
            DgvTeam.AllowUserToResizeRows = False
            DgvTeam.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            DgvTeam.TabIndex = 0
            Dim TBL As dseTransaksiStockTaking.tbltrans_stock_taking_staffDataTable
            Dim row As dseTransaksiStockTaking.tbltrans_stock_taking_staffRow
            TBL = DTAStaff.GetDataBy_TeamIdentifier(RowTeam.Identifier)
            If TBL.Rows.Count > 0 Then
                Dim i As Int16 = 1
                For Each row In TBL.Rows
                    DgvTeam.Rows.Add(i, row.Nama, row.Remark)
                    i += 1
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAStaff.Dispose()
        End Try
    End Sub

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_getNoTransaksiByPrimary(vPrimary)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_getNoTransaksiByPrimary(vPrimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function getJmlByNoTransaksi(ByVal vNoTransaksi As String) As Integer
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi).HasValue Then
                Return DTA.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsikanListTagNo(ByVal vlistItem As List(Of String), ByVal vPrimaryTaking As Double, ByVal vIdentifier As String) As Boolean
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking_tagNTeamTableAdapter
        Dim TBL As dseTransaksiStockTaking.tbltrans_stock_taking_tagNTeamDataTable
        Dim row As dseTransaksiStockTaking.tbltrans_stock_taking_tagNTeamRow
        Try
            vlistItem.Clear()
            TBL = DTA.GetDataBy_PrimaryTakingNIdentifier(vIdentifier, vPrimaryTaking)
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    vlistItem.Add(row.Tag_No)
                Next
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsikanTable(ByVal vTBL As dseTransaksiStockTaking.tbltrans_stock_taking_team1DataTable, ByVal vPrimaryTaking As Double) As Boolean
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking_team1TableAdapter
        Try
            vTBL.Rows.Clear()
            DTA.Fill(vTBL, CStr(vPrimaryTaking))
            If vTBL.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanBaru(ByVal vUserCode As String, ByVal TabCont As TabControl, _
    Optional ByVal vCekDataBenar As Boolean = True, Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTATeam As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking_teamTableAdapter
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Dim DTAStaff As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking_staffTableAdapter
        Dim ObjClosing As New ClsTransaksiMonthlyClosingItem
        Dim ObjCD As New Core_Module.ClsCekData

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now

        Try
            If vCekDataBenar = True Then
                'cek kebenaran data
                '1. tidak boleh ada stock taking pada bulan yang sudah diclosing
                If ObjCD.GetLasDateOfTheMonth(ObjClosing.GetLastClosingPeriod) >= Me.mStartDate Then
                    MsgBox("There closing for " + Format(Me.StartDate, "MMM yyyy") + ", you cannot save this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                '2. dalam 1 bulan hanya boleh ada 1 data stock taking
                If Me.GetJml(ObjCD.GetFirstDateofTheMonth(Me.mStartDate), ObjCD.GetLasDateOfTheMonth(Me.mStartDate)) > 0 Then
                    MsgBox("There stock taking for " + Format(Me.StartDate, "MMM yyyy") + ", you cannot save this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If vFromCutOff = True Then
                CreatedBy_ = Me.CreatedBy
                CreatedDate_ = Me.CreatedDate
                ModifiedBy_ = Me.ModifiedBy
                ModifiedDate_ = Me.ModifiedDate
                TimeStamp_ = Me.TimeStamp
            End If

            'simpan data header
            If DTA.Insert(Me.NoTransaksi, Me.StartDate, Me.TotalTeam, Me.Desc, Me.Status, Me.ProcessDate, CreatedBy_, CreatedDate_, ModifiedBy_, ModifiedDate_, TimeStamp_) > 0 Then
                Me.Primary = DTA.ScalarQuery_getPrimaryByNoTransaksi(Me.NoTransaksi)
                'simpan data item team
                Dim i As TabPage
                Dim TeamJobDesc As String = ""
                Dim TeamNo As String = ""
                Dim cont As Control
                Dim j As Integer
                For j = 1 To TabCont.TabPages.Count - 1
                    i = TabCont.TabPages(j)
                    For Each cont In i.Controls
                        'cari team jobdesc
                        If cont.Name.StartsWith("txtTeamJobDesc") Then
                            TeamJobDesc = cont.Text
                            Exit For
                        End If
                    Next
                    For Each cont In i.Controls
                        'cari no team
                        If cont.Name.StartsWith("txtTeamNumber") Then
                            TeamNo = cont.Text
                            Exit For
                        End If
                    Next
                    'simpan team
                    Dim Identifier As String = CStr(Me.Primary) + "-" + CStr(j)
                    DTATeam.Insert(Me.Primary, Identifier, TeamNo, TeamJobDesc)
                    'cai control di dgv kemudian simpan staff
                    For Each cont In i.Controls
                        If cont.Name.StartsWith("gbo") Then
                            Dim cont1 As Control
                            For Each cont1 In cont.Controls
                                If cont1.Name.StartsWith("DataGridView") Then
                                    Dim Dgv As DataGridView = TryCast(cont1, DataGridView)
                                    Dim row As DataGridViewRow
                                    For Each row In Dgv.Rows
                                        'simpan data staff
                                        If row.Cells(1).Value <> "" Or row.Cells(2).Value <> "" Then
                                            DTAStaff.Insert(row.Cells(1).Value, row.Cells(2).Value, Identifier)
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next
                Next
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
            DTATeam.Dispose()
            DTAStaff.Dispose()
            ObjClosing = Nothing
            ObjCD = Nothing
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vUserCode As String, ByVal TabCont As TabControl) As Boolean
        Dim DTATeam As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking_teamTableAdapter
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Dim DTAStaff As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking_staffTableAdapter
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjClosing As New ClsTransaksiMonthlyClosingItem

        Try
            'cek kebenaran data
            If Me.ApakahBolehEdit(Me.Primary) = False Then
                Exit Try
            End If
            '1. tidak boleh ada stock taking pada bulan yang sudah diclosing
            If ObjCD.GetLasDateOfTheMonth(ObjClosing.GetLastClosingPeriod) >= Me.mStartDate Then
                MsgBox("There closing for " + Format(Me.StartDate, "MMM yyyy") + ", you cannot save this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            '2. dalam 1 bulan hanya boleh ada 1 data stock taking
            If Me.GetJml(ObjCD.GetFirstDateofTheMonth(Me.mStartDate), ObjCD.GetLasDateOfTheMonth(Me.mStartDate), Me.mPrimary) > 0 Then
                MsgBox("There stock taking for " + Format(Me.StartDate, "MMM yyyy") + ", you cannot save this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'simpan data header
            If DTA.Update(Me.NoTransaksi, Me.StartDate, Me.TotalTeam, Me.Desc, Me.Status, Me.ProcessDate, vUserCode, Now, Me.TimeStamp, Me.Primary) > 0 Then
                Me.Primary = DTA.ScalarQuery_getPrimaryByNoTransaksi(Me.NoTransaksi)
                Me.TimeStamp = DTA.ScalarQuery_getTimestamp(CStr(Me.Primary))
                Me.ModifiedDate = Me.TimeStamp
                Me.ModifiedBy = vUserCode
                'hapus data item team, ini akan otomatis menghapus staff
                DTATeam.Delete(Me.Primary)
                'Me.Hapus(Me.Primary)
                'simpan data item team            
                'simpan data item team member
                'simpan data item team
                Dim i As TabPage
                Dim TeamJobDesc As String = ""
                Dim TeamNo As String = ""
                Dim cont As Control
                Dim j As Integer
                For j = 1 To TabCont.TabPages.Count - 1
                    i = TabCont.TabPages(j)
                    For Each cont In i.Controls
                        'cari team jobdesc
                        If cont.Name.StartsWith("txtTeamJobDesc") Then
                            TeamJobDesc = cont.Text
                            Exit For
                        End If
                    Next
                    For Each cont In i.Controls
                        'cari no team
                        If cont.Name.StartsWith("txtTeamNumber") Then
                            TeamNo = cont.Text
                            Exit For
                        End If
                    Next
                    'simpan team
                    Dim Identifier As String = CStr(Me.Primary) + "-" + CStr(j)
                    DTATeam.Insert(Me.Primary, Identifier, TeamNo, TeamJobDesc)
                    'cai control di dgv kemudian simpan staff
                    For Each cont In i.Controls
                        If cont.Name.StartsWith("gbo") Then
                            Dim cont1 As Control
                            For Each cont1 In cont.Controls
                                If cont1.Name.StartsWith("DataGridView") Then
                                    Dim Dgv As DataGridView = TryCast(cont1, DataGridView)
                                    Dim row As DataGridViewRow
                                    For Each row In Dgv.Rows
                                        'simpan data staff
                                        If row.Cells(1).Value <> "" Or row.Cells(2).Value <> "" Then
                                            DTAStaff.Insert(row.Cells(1).Value, row.Cells(2).Value, Identifier)
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next
                Next
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTATeam.Dispose()
            DTA.Dispose()
            DTAStaff.Dispose()
            ObjCD = Nothing
            ObjClosing = Nothing
        End Try
    End Function

    Public Sub HapusDataPenyesuaianStockTaking(ByVal vPrimaryClosing As Double)
        Dim DTA As New DseLaporanStockTakingBaruTableAdapters.tbltrans_stock_taking_laporancostTableAdapter

        Try
            DTA.Delete(vPrimaryClosing)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim DTATeam As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_taking_teamTableAdapter
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Try
            'cek apakah boleh hapus
            If Me.ApakahBolehHapus(Me.Primary) = False Then
                Exit Try
            End If
            'hapus data item
            DTA.Delete(Me.Primary)
            'hapus semua data transaksi jika ada
            DTATeam.Delete(Me.Primary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTATeam.Dispose()
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiItemStartForFormPrinting(ByVal TBLStock As dseLaporanStockTaking.tblmaster_productqtyDataTable)
        Dim DTAStock As New dseLaporanStockTakingTableAdapters.tblmaster_productqtyTableAdapter
        Dim PrimaryStockTaking As Double = Me.GetPrimaryMax

        Try
            DTAStock.Connection.ConnectionString = Core_Module.GetConstring
            'bentuk kembali kondisi stock seperti semula ketika melakukan stock opname, stock yang
            'keluar pada hari stock opname dianggap keluar setelah dilakukan stock opname
            DTAStock.Fill(TBLStock)
            Me.GetItemStart(TBLStock, PrimaryStockTaking, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAStock.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' mengisikan item buffer stock taking
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub IsiItemBuffer()
        Dim DTAItemBuffer As New dseLaporanStockTakingTableAdapters.tbltrans_stock_taking_itembufferTableAdapter
        Dim TBLBuffer As New dseLaporanStockTaking.tbltrans_stock_taking_itembufferDataTable
        Dim rowBuffer As dseLaporanStockTaking.tbltrans_stock_taking_itembufferRow

        Dim TBLStock As New dseLaporanStockTaking.tblmaster_productqtyDataTable
        Dim rowStock As dseLaporanStockTaking.tblmaster_productqtyRow
        Dim DTAStock As New dseLaporanStockTakingTableAdapters.tblmaster_productqtyTableAdapter

        Dim TBLTaking As New dseLaporanStockTaking.StockTagItem1DataTable
        Dim DTATaking As New dseLaporanStockTakingTableAdapters.StockTagItem1TableAdapter
        Dim rowTaking As dseLaporanStockTaking.StockTagItem1Row

        Dim PrimaryStockTaking As Double = Me.GetPrimaryMax
        Dim Filter As String = ""

        Try
            'kosongkan item buffer yang sudah ada
            DTAItemBuffer.Delete()

            'bentuk kembali kondisi stock seperti semula ketika melakukan stock opname, stock yang
            'keluar pada hari stock opname dianggap keluar setelah dilakukan stock opname
            DTAStock.Fill(TBLStock)
            Me.GetItemStart(TBLStock, PrimaryStockTaking)

            'tambahkan item dari stock on hand - item start
            For Each rowStock In TBLStock.Rows
                rowBuffer = TBLBuffer.Newtbltrans_stock_taking_itembufferRow
                With rowBuffer
                    .Difference = 0
                    .Expiry_Date = rowStock.Expiry_Date
                    .Lot_No = rowStock.Lot_No
                    .Primary = TBLBuffer.Rows.Count + 1
                    .Primary_Product = rowStock.Primary_Product
                    .Primary_Storage_Area = rowStock.Primary_Storage_Area
                    .Qty_On_Hand = rowStock.Qty
                    .Qty_Secondary = rowStock.Qty_Secondary
                    .Qty_Taking = 0
                End With
                TBLBuffer.Rows.Add(rowBuffer)
            Next

            'untuk setiap item stock tag, jika ada maka masukkan nilai qty taking
            DTATaking.Fill(TBLTaking, PrimaryStockTaking)
            For Each rowTaking In TBLTaking.Rows
                'filter dim tabel buffer apakah sudah ada data yang dimaksud
                'karens 1 lot untuk 1 expiry, maka filter expiry date tidak digunakan 
                'tadinya filter menggunakan expiry date, tetapi kemudian bermasalah kalo expiry date pada check tag
                'berbeda dengan expiry date yang digunakan sekarang, dan qty on hand > 0 saat taking dan qty onhand =0 saat proses
                Filter = "[Lot No] = '" + rowTaking.Lot_No + "' AND [Primary Product] =" + rowTaking.Primary_Product.ToString + _
                " AND [Qty Secondary] = " + rowTaking.Qty_Secondary.ToString

                If rowTaking.Location_Controlled = 1 Then
                    Filter = Filter + " AND [Primary Storage Area] = " + rowTaking.Primary_Storage_Area.ToString
                Else
                    Filter = Filter + " AND [Primary Storage Area] = -1"
                End If

                TBLBuffer.DefaultView.RowFilter = Filter
                If TBLBuffer.DefaultView.Count > 0 Then
                    'jika ada datanya maka tambahkan qty taking
                    rowBuffer = TBLBuffer.DefaultView.Item(0).Row
                    rowBuffer.Qty_Taking += rowTaking.Qty
                    'batch number di update sesuai dengan data taking
                    rowBuffer.Expiry_Date = rowTaking.Expiry_Date
                Else
                    'bila tidak ada datanya maka tambahkan baris baru
                    rowBuffer = TBLBuffer.Newtbltrans_stock_taking_itembufferRow
                    With rowBuffer
                        .Difference = 0
                        .Expiry_Date = rowTaking.Expiry_Date
                        .Lot_No = rowTaking.Lot_No
                        .Primary = TBLBuffer.Rows.Count + 1
                        .Primary_Product = rowTaking.Primary_Product
                        If rowTaking.Location_Controlled = 1 Then
                            .Primary_Storage_Area = rowTaking.Primary_Storage_Area
                        Else
                            .Primary_Storage_Area = -1
                        End If
                        .Qty_Taking = rowTaking.Qty
                        .Qty_Secondary = rowTaking.Qty_Secondary
                        .Qty_On_Hand = 0
                    End With

                    TBLBuffer.Rows.Add(rowBuffer)
                End If
            Next

            'simpan perubahan ke dalam basis data
            For Each rowBuffer In TBLBuffer.Rows
                With rowBuffer
                    'hitung difference
                    rowBuffer.Difference = .Qty_Taking - .Qty_On_Hand
                    'simpan ke dalam basis data, hanya yang berbeda saja
                    'DTAItemBuffer.Insert(.Primary_Product, .Qty_Secondary, .Lot_No, .Expiry_Date, .Primary_Storage_Area, _
                    '.Qty_On_Hand, .Qty_Taking, .Difference)
                End With
            Next

            Me.Simpan_ItemBuffer_UsingScript(TBLBuffer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemBuffer.Dispose()
            TBLBuffer.Dispose()
            TBLStock.Dispose()
            DTAStock.Dispose()
            TBLTaking.Dispose()
            DTATaking.Dispose()
        End Try
    End Sub

    Public Sub SimpanDataPenyesuaian_StockTaking(ByVal vPrimaryClosing As Double, ByVal vPrimaryStockTaking As Double)
        Dim DTA As New DseLaporanStockTakingBaruTableAdapters.tbltrans_stock_taking_laporancostTableAdapter
        Dim TBL As New DseLaporanStockTakingBaru.tbltrans_stock_taking_laporancostDataTable
        Dim row As DseLaporanStockTakingBaru.tbltrans_stock_taking_laporancostRow
        Dim DTA1 As New DseLaporanStockTakingBaruTableAdapters.DataTable3TableAdapter
        Dim TBL1 As New DseLaporanStockTakingBaru.DataTable3DataTable
        Dim row1 As DseLaporanStockTakingBaru.DataTable3Row
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust
        Dim TipeTransaksiAsal1 As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust
        Dim TipeItemDetail As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
        Dim TipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
        Dim Amount As Decimal = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            'kumpulkan data stock taking yang ada pada periode yang bersangkutan
            DTA.FillBy_PrimaryTaking(TBL, vPrimaryStockTaking)

            'kumpulkan item cost
            DTA1.Fill(TBL1, vPrimaryClosing, TipeTransaksiAsal, TipeTransaksiAsal1, TipeItemDetail, TipeItemDetail1)

            'untuk setiap data yang ada, hitung amountnya
            For Each row In TBL.Rows
                row.primary_closing = vPrimaryClosing
                If row.difference <> 0 Then
                    TBL1.DefaultView.RowFilter = "[primary code no] = " + row.primary_product.ToString
                    Amount = 0

                    For Each dv As Data.DataRowView In TBL1.DefaultView
                        row1 = dv.Row
                        Amount += row1.amount
                    Next

                    If row.difference > 0 Then
                        row.sub_total = Amount
                    Else
                        row.sub_total = Amount * -1
                    End If
                End If
            Next

            'hapuskan data lama
            DTA.DeleteQuery_ByPrimaryTaking(vPrimaryStockTaking)

            'simpan data baru
            Me.Simpan_ItemCost_UsingScript(TBL, vPrimaryStockTaking, vPrimaryClosing)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Sub

    Private Function Simpan_ItemCost_UsingScript(ByVal vTBL As DseLaporanStockTakingBaru.tbltrans_stock_taking_laporancostDataTable, _
    ByVal vPrimaryStockTaking As Double, ByVal vPrimaryClosing As Double) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemCost(StartRowIndex, BatchQty, vTBL, vPrimaryStockTaking, vPrimaryClosing)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemCost(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseLaporanStockTakingBaru.tbltrans_stock_taking_laporancostDataTable, ByVal vPrimaryStockTaking As Double, ByVal vPrimaryClosing As Double) As String
        Dim retval As String = ""
        Dim row As DseLaporanStockTakingBaru.tbltrans_stock_taking_laporancostRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_stock_taking_laporancost` (`primary closing`, `primary stock taking`, `primary product`, `qty on hand`, `qty taking`, `difference`, `primary unit`, `sub total`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + vPrimaryClosing.ToString + ", " + vPrimaryStockTaking.ToString + ", " + _
                        CDbl(.primary_product).ToString + "," + CDec(.qty_on_hand).ToString + "," _
                        + CDec(.qty_taking).ToString + "," _
                        + CDec(.difference).ToString + "," _
                        + CDbl(.primary_unit).ToString + "," _
                        + CDec(.sub_total).ToString + ")" + Delimiter
                    End With

                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetPrimary_ByPeriode(ByVal vStartDate As Date, ByVal vEndDate As Date) As Double
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tbltrans_stock_takingTableAdapter
        Dim TBL As New dseTransaksiStockTaking.tbltrans_stock_takingDataTable
        Dim row As dseTransaksiStockTaking.tbltrans_stock_takingRow

        Try
            DTA.FillBy_TanggalDariSampai(TBL, vEndDate, vStartDate)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Return row.Primary
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Private Sub GetItemStart(ByVal TBL As dseLaporanStockTaking.tblmaster_productqtyDataTable, _
    ByVal vPrimaryStockTaking As Double, Optional ByVal SimpanDB As Boolean = True)
        Dim DTA As New dseLaporanStockTakingTableAdapters.tbltrans_stock_taking_itemstart1TableAdapter

        Dim DTAInOut As New dseTransaksiStockInOutTableAdapters.tbltrans_stockinout_item_view0TableAdapter
        Dim TBLInOut As New dseTransaksiStockInOut.tbltrans_stockinout_item_view0DataTable
        Dim RowInOut As dseTransaksiStockInOut.tbltrans_stockinout_item_view0Row
        Dim TipeInOut As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust

        Dim DTABefore As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
        Dim TBLBefore As New dseTransaksiStockBundling.tbltransaksi_stockbundling_item_beforeDataTable
        Dim rowBefore As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_beforeRow
        Dim DTAAfter As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim TBLAfter As New dseTransaksiStockBundling.tbltransaksi_stockbundling_item_afterDataTable
        Dim rowAfter As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_afterRow

        Dim DTATransfer As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransfer_itemTableAdapter
        Dim TBLTransfer As New dseTransaksiStockTransfer.tbltransaksi_stocktransfer_itemDataTable
        Dim rowTransfer As dseTransaksiStockTransfer.tbltransaksi_stocktransfer_itemRow

        Dim StartDate As Date = Me.GetStartDate(vPrimaryStockTaking)
        'Dim rowStart As dseLaporanStockTaking.tblmaster_productqtyRow
        Dim TipeProductGoods As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods

        Try
            'start date termasuk, jadi semua transaksi stock pada tanggal stock taking dianggap dilakukan
            'setelah proses stock taking
            'hanya memperhitungkan item2 goods
            'stock in out
            DTAInOut.FillBy_TglDariSampai(TBLInOut, #1/1/2100#, StartDate, CInt(TipeProductGoods))
            For Each RowInOut In TBLInOut.Rows
                TipeInOut = RowInOut.Tipe_Transaksi
                With RowInOut
                    If TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
                    TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess Or _
                    TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process Or _
                    TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
                    TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Then
                        'kurangkan item
                        Me.KurangkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, .Batch_No, _
                        .Expiry_Date, .Primary_Storage_Area)
                    ElseIf TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Or _
                        TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Process Or _
                        TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Or _
                        TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                        'tambahkan item
                        Me.TambahkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, .Batch_No, _
                        .Expiry_Date, .Primary_Storage_Area)
                    End If
                End With
            Next

            'stock bundling -> pasti goods
            DTABefore.FillBy_TanggalSampaiDari2(TBLBefore, #1/1/2100#, StartDate, CInt(TipeProductGoods))
            For Each rowBefore In TBLBefore.Rows
                With rowBefore
                    Me.TambahkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, _
                    .Batch_No, .Expiry_Date, .Primary_Storage_Area)
                End With
            Next
            DTAAfter.FillBy_TanggalSampaiDari2(TBLAfter, #1/1/2100#, StartDate, CInt(TipeProductGoods))
            For Each rowAfter In TBLAfter.Rows
                With rowAfter
                    Me.KurangkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, _
                    .Batch_No, .Expiry_Date, .Primary_Storage_Area)
                End With
            Next

            'stock transfer-pasti goods, jadi tidak dibatasi lagi.
            DTATransfer.FillBy_TanggalSampaiDari(TBLTransfer, #1/1/2100#, StartDate)
            For Each rowTransfer In TBLTransfer.Rows
                With rowTransfer
                    'kurangkan jumlah dari lokasi tujuan
                    Me.KurangkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, _
                    .Batch_No, .Expiry_Date, .primary_destination_area)
                    'tambahkan jumlah dari lokasi asal
                    Me.TambahkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, _
                    .Batch_No, .Expiry_Date, .Primary_Storage_Area)
                End With
            Next

            'buang yang tipe productnya jasa

            'simpan sebagai item start
            If SimpanDB = True Then
                DTA.Delete()
                'For Each rowStart In TBL.Rows
                '    With rowStart
                '        DTA.Insert(.Primary_Product, .Qty, .Qty_Secondary, .Lot_No, _
                '        .Expiry_Date, .Primary_Storage_Area)
                '    End With
                'Next
                Me.Simpan_ItemStart_UsingScript(TBL)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAInOut.Dispose()
            TBLInOut.Dispose()
            DTABefore.Dispose()
            TBLBefore.Dispose()
            DTAAfter.Dispose()
            TBLAfter.Dispose()
            DTATransfer.Dispose()
            TBLTransfer.Dispose()
        End Try
    End Sub

    Private Function Simpan_ItemBuffer_UsingScript(ByVal vTBL As dseLaporanStockTaking.tbltrans_stock_taking_itembufferDataTable) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemBuffer(StartRowIndex, BatchQty, vTBL)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemBuffer(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vTBL As dseLaporanStockTaking.tbltrans_stock_taking_itembufferDataTable) As String
        Dim retval As String = ""
        Dim row As dseLaporanStockTaking.tbltrans_stock_taking_itembufferRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_stock_taking_itembuffer` (`Primary Product`, `Qty Secondary`, `Lot No`, `Expiry Date`, `Primary Storage Area`, `Qty On Hand`, `Qty Taking`, `Difference`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        If .Lot_No <> "" Then
                            If .Lot_No.Contains("\") = True Then
                                .Lot_No = .Lot_No.Replace("\", "\\")
                            End If
                            If .Lot_No.Contains(vbCrLf) = True Then
                                .Lot_No = .Lot_No.Replace(vbCrLf, "\r\n")
                            End If
                            If .Lot_No.Contains("'") = True Then
                                .Lot_No = .Lot_No.Replace("'", "''")
                            End If
                        End If

                        retval = retval + vbCrLf + "(" + CDbl(.Primary_Product).ToString + "," + CDec(.Qty_Secondary).ToString + "," _
                        + "'" + .Lot_No + "'," _
                        + "'" + Format(CDate(.Expiry_Date), "yyyy-MM-dd") + "'," _
                        + CDbl(.Primary_Storage_Area).ToString + "," _
                        + CDec(.Qty_On_Hand).ToString + "," _
                        + CDec(.Qty_Taking).ToString + "," _
                        + CDec(.Difference).ToString + "" _
                        + ")" + Delimiter
                    End With

                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Function Simpan_ItemStart_UsingScript(ByVal vTBL As dseLaporanStockTaking.tblmaster_productqtyDataTable) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemStart(StartRowIndex, BatchQty, vTBL)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemStart(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vTBL As dseLaporanStockTaking.tblmaster_productqtyDataTable) As String
        Dim retval As String = ""
        Dim row As dseLaporanStockTaking.tblmaster_productqtyRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_stock_taking_itemstart` (`Primary Product`, `Qty`, `Qty Secondary`, `Lot No`, `Expiry Date`, `Primary Storage Area`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        If .Lot_No <> "" Then
                            If .Lot_No.Contains("\") = True Then
                                .Lot_No = .Lot_No.Replace("\", "\\")
                            End If
                            If .Lot_No.Contains(vbCrLf) = True Then
                                .Lot_No = .Lot_No.Replace(vbCrLf, "\r\n")
                            End If
                            If .Lot_No.Contains("'") = True Then
                                .Lot_No = .Lot_No.Replace("'", "''")
                            End If
                        End If

                        retval = retval + vbCrLf + "(" + CDbl(.Primary_Product).ToString + "," + CDec(.Qty).ToString + "," _
                        + CDec(.Qty_Secondary).ToString + "," _
                        + "'" + .Lot_No + "'," _
                        + "'" + Format(CDate(.Expiry_Date), "yyyy-MM-dd") + "'," _
                        + CDbl(.Primary_Storage_Area).ToString + "" _
                        + ")" + Delimiter
                    End With

                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Sub TambahkanItem(ByVal TBL As dseLaporanStockTaking.tblmaster_productqtyDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vQty As Double, ByVal vQtySecondary As Double, _
    ByVal vLotNo As String, ByVal vExpiryDate As Date, ByVal vPrimaryStorageArea As Double)
        Dim filter As String = ""
        Dim rowInOut As dseLaporanStockTaking.tblmaster_productqtyRow

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + _
            " AND [Qty Secondary] = " + vQtySecondary.ToString + " AND [Primary Storage Area] = " + _
            vPrimaryStorageArea.ToString + " AND [Expiry Date] = #" + vExpiryDate.ToShortDateString + _
            "# AND [Lot No] = '" + vLotNo + "'"

            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count = 0 Then
                'tambahkan baris baru
                If TBL.Rows.Count > 0 Then
                    rowInOut = TBL.Rows(TBL.Rows.Count - 1)
                    filter = rowInOut.Primary
                Else
                    filter = "0"
                End If

                rowInOut = TBL.Newtblmaster_productqtyRow
                With rowInOut
                    .Expiry_Date = vExpiryDate
                    If vPrimaryStorageArea <> -1 Then
                        .Location_Controlled = 1
                    Else
                        .Location_Controlled = 0
                    End If
                    .Lot_No = vLotNo
                    .Primary = CDbl(filter) + 1
                    .Primary_Product = vPrimaryProduct
                    .Primary_Storage_Area = vPrimaryStorageArea
                    .Qty = vQty
                    .Qty_Secondary = vQtySecondary
                End With
                TBL.Rows.Add(rowInOut)
            Else
                'tambahkan nilai qty
                rowInOut = TBL.DefaultView(0).Row
                rowInOut.Qty += vQty
                If rowInOut.Qty = 0 Then
                    TBL.Rows.Remove(rowInOut)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KurangkanItem(ByVal TBL As dseLaporanStockTaking.tblmaster_productqtyDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vQty As Double, ByVal vQtySecondary As Double, _
    ByVal vLotNo As String, ByVal vExpiryDate As Date, ByVal vPrimaryStorageArea As Double)
        Dim filter As String = ""
        Dim rowInOut As dseLaporanStockTaking.tblmaster_productqtyRow

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + _
            " AND [Qty Secondary] = " + vQtySecondary.ToString + " AND [Primary Storage Area] = " + _
            vPrimaryStorageArea.ToString + " AND [Expiry Date] = #" + vExpiryDate.ToShortDateString + _
            "# AND [Lot No] = '" + vLotNo + "'"

            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count = 0 Then
                'tambahkan baris baru
                rowInOut = TBL.Rows(TBL.Rows.Count - 1)
                filter = rowInOut.Primary
                rowInOut = TBL.Newtblmaster_productqtyRow
                With rowInOut
                    .Expiry_Date = vExpiryDate
                    If vPrimaryStorageArea <> -1 Then
                        .Location_Controlled = 1
                    Else
                        .Location_Controlled = 0
                    End If
                    .Lot_No = vLotNo
                    .Primary = CDbl(filter) + 1
                    .Primary_Product = vPrimaryProduct
                    .Primary_Storage_Area = vPrimaryStorageArea
                    .Qty = vQty * (-1)
                    .Qty_Secondary = vQtySecondary
                End With
                TBL.Rows.Add(rowInOut)
            Else
                'kurangkan nilai qty
                rowInOut = TBL.DefaultView(0).Row
                rowInOut.Qty -= vQty
                If rowInOut.Qty = 0 Then
                    TBL.Rows.Remove(rowInOut)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
