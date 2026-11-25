Option Explicit On
Option Strict Off

Public Class ClsMasterBudget
    Private mPrimary As Double = -1
    Private mFinancialYear As Date = #1/1/1950#
    Private mCreatedBy As String = ""
    Private mCreatedDate As Date = Now
    Private mModifiedBy As String = ""
    Private mModifiedDate As Date = Now
    Private mTimestamp As Date = Now
    Private mTipeBudget As EnumTipeBudget = EnumTipeBudget.BS

    'komponen yang berkaitan dengan realisasi budget
    Dim ObjCD As New Core_Module.ClsCekData
    Dim DTA_Beginning_Balance As New DseLaporanTrialBalanceTableAdapters.tbl_beginning_balanceTableAdapter
    Dim TBL_Beginning_Balance As New DseLaporanTrialBalance.tbl_beginning_balanceDataTable
    Dim DTA_TrialBalance As New DseLaporanTrialBalanceTableAdapters.Trial_BalanceTableAdapter

    Public Enum EnumTipeBudget As Integer
        BS
        PL
    End Enum

    Public Property TipeBudget() As EnumTipeBudget
        Get
            Return Me.mTipeBudget
        End Get
        Set(ByVal value As EnumTipeBudget)
            mTipeBudget = value
        End Set
    End Property

    Public Property FinancialYear() As Date
        Get
            Return Me.mFinancialYear
        End Get
        Set(ByVal value As Date)
            mFinancialYear = value
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
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
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

    Public Property ModifiedDate() As Date
        Get
            Return mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property Timestamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
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

    Public Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vTipeBudget As EnumTipeBudget, Optional ByVal vTampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseMasterBudgetTableAdapters.tblmaster_budgetTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'data salah jika
            '1. untuk data baru, jika sudah ada data budget untuk financial year tsb
            If vBaru = True Then
                If DTA.ScalarQuery_GetJml2(Me.mFinancialYear, vTipeBudget).HasValue = True Then
                    If DTA.ScalarQuery_GetJml2(Me.mFinancialYear, vTipeBudget).Value > 0 Then
                        If vTampilkanPesan = True Then
                            MsgBox("There is budget already for " + Format(Me.FinancialYear, "yyyy") + _
                            ".", MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanBaru(ByVal vDGV As DataGridView, ByVal ij As FrmMasterBudget, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterBudgetTableAdapters.tblmaster_budgetTableAdapter
        Dim DTA_Item As New DseMasterBudgetTableAdapters.tblmaster_budget_itemTableAdapter
        Dim rowDGV As DataGridViewRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Item.Connection.ConnectionString = DTA.Connection.ConnectionString

            If Me.ApakahDataBenar(True, Me.mTipeBudget) = True Then
                'hitung level otorisasi yang dibutuhkan
                'insert query
                If DTA.Insert(Me.mFinancialYear, vUserCode, Now, vUserCode, Now, Now, Me.mTipeBudget) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mFinancialYear, Me.mTipeBudget).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    'simpan data item baru
                    For Each rowDGV In vDGV.Rows
                        With rowDGV
                            DTA_Item.Insert(Me.mPrimary, .Cells(ij.ColPrimaryAkun.Index).Value, .Cells(ij.ColJan.Index).Value, _
                            .Cells(ij.ColFeb.Index).Value, .Cells(ij.ColMar.Index).Value, .Cells(ij.ColApr.Index).Value, _
                            .Cells(ij.ColMay.Index).Value, .Cells(ij.ColJune.Index).Value, .Cells(ij.ColJul.Index).Value, _
                            .Cells(ij.ColAgt.Index).Value, .Cells(ij.ColSept.Index).Value, .Cells(ij.ColOct.Index).Value, _
                            .Cells(ij.ColNov.Index).Value, .Cells(ij.ColDec.Index).Value, .Cells(ij.ColRemarks.Index).Value)
                        End With
                    Next

                    'sampe sini return true
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Format(Me.FinancialYear, "yyyy-MM-dd") + "' for key 2" Then
                Return Me.SimpanBaru(vDGV, ij, vUserCode)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            DTA_Item.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vDGV As DataGridView, ByVal ij As FrmMasterBudget, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterBudgetTableAdapters.tblmaster_budgetTableAdapter
        Dim DTA_Item As New DseMasterBudgetTableAdapters.tblmaster_budget_itemTableAdapter
        Dim rowDGV As DataGridViewRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Item.Connection.ConnectionString = DTA.Connection.ConnectionString

            If Me.ApakahDataBenar(False, Me.mTipeBudget) = True Then
                'hitung level otorisasi yang dibutuhkan
                'insert query
                If DTA.Update(Me.mFinancialYear, vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah update query
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    'hapus data lama
                    DTA_Item.Delete(Me.mPrimary)
                    'simpan data item baru
                    For Each rowDGV In vDGV.Rows
                        With rowDGV
                            DTA_Item.Insert(Me.mPrimary, .Cells(ij.ColPrimaryAkun.Index).Value, .Cells(ij.ColJan.Index).Value, _
                            .Cells(ij.ColFeb.Index).Value, .Cells(ij.ColMar.Index).Value, .Cells(ij.ColApr.Index).Value, _
                            .Cells(ij.ColMay.Index).Value, .Cells(ij.ColJune.Index).Value, .Cells(ij.ColJul.Index).Value, _
                            .Cells(ij.ColAgt.Index).Value, .Cells(ij.ColSept.Index).Value, .Cells(ij.ColOct.Index).Value, _
                            .Cells(ij.ColNov.Index).Value, .Cells(ij.ColDec.Index).Value, .Cells(ij.ColRemarks.Index).Value)
                        End With
                    Next

                    'sampe sini return true
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Format(Me.FinancialYear, "yyyy-MM-dd") + "' for key 2" Then
                Return Me.SimpanPerubahan(vDGV, ij, vUserCode)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            DTA_Item.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal vTampilkanPesan As Boolean, _
    ByVal vStrictUser As Boolean) As Boolean
        Dim DTA As New DseMasterBudgetTableAdapters.tblmaster_budgetTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUserCode, vTampilkanPesan, True) = True Then
                DTA.Delete(vPrimary)
                'retval
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal vTampilkanPesan As Boolean, _
    Optional ByVal vStrictUser As Boolean = True) As Boolean
        Dim DTA As New DseMasterBudgetTableAdapters.tblmaster_budgetTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'tidak boleh edit jika bukan user yang bersangkutan
            Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimary)
            If CreatedBy = vUserCode Then
            Else
                If vTampilkanPesan = True Then
                    MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal vTampilkanPesan As Boolean, Optional ByVal vStrictUser As Boolean = True) As Boolean
        Dim DTA As New DseMasterBudgetTableAdapters.tblmaster_budgetTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'tidak boleh edit jika bukan user yang bersangkutan
            Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimary)
            If CreatedBy = vUserCode Then
            Else
                If vTampilkanPesan = True Then
                    MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' fungsi ini mengembalikan budget dalam bentuk trial balance
    ''' Tabel yang terdiri dari Primary - Account Code - Saldo
    ''' </summary>
    ''' <param name="Period"></param>
    ''' <remarks></remarks>
    Public Sub GetBudget_BS(ByVal Period As Date, ByVal TBL As DseLaporanTrialBalance.Trial_BalanceDataTable)
        Dim vPrimaryBudget As Double = -1
        Dim DTA As New DseMasterBudgetTableAdapters.DataTable1TableAdapter
        Dim TBLItem As New DseMasterBudget.DataTable1DataTable
        Dim row As DseMasterBudget.DataTable1Row
        Dim rowBudget As DseLaporanTrialBalance.Trial_BalanceRow
        Dim Val As Double = 0
        Dim SN As ClsMasterAkun.EnumSaldoNormal = ClsMasterAkun.EnumSaldoNormal.Debet
        Dim ObjCD As New Core_Module.ClsCekData
        DTA.ClearBeforeFill = False

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBLItem.Rows.Clear()


            vPrimaryBudget = Me.GetPrimary(ObjCD.GetFisrtDateOfTheYear(Period), FrmMasterBudget.EnumFormatYearly.Balance_Sheet)
            DTA.Fill(TBLItem, vPrimaryBudget)
            'Else
            'vPrimaryBudget = Me.GetPrimary(ObjCD.GetFisrtDateOfTheYear(Period), FrmMasterBudget.EnumFormatYearly.Profit_N_Loss)
            'DTA.Fill(TBLItem, vPrimaryBudget)
            'End If

            For Each row In TBLItem.Rows
                Val = 0
                rowBudget = TBL.NewTrial_BalanceRow
                With rowBudget
                    .Primary_Akun = row.Primary_Akun
                    .Saldo_Normal = row.Saldo_Normal
                    .Debet = 0
                    .Credit = 0
                    .Akun = ""
                    .Header = row.Header
                    .Level_Perkiraan = row.Level_Perkiraan
                    .Beginning_Balance = row.Beginning_Balance
                    .Saldo = 0
                    .Aktif = row.Aktif
                    .Description = row.Description
                    .Tipe_Perkiraan = row.Tipe_Perkiraan
                    .Urutan = row.Urutan

                    .Account_Code = row.Account_Code

                    Val = 0
                    If Period.Month = 1 Then
                        Val += row.Januari
                    End If
                    If Period.Month = 2 Then
                        Val += row.Feb
                    End If
                    If Period.Month = 3 Then
                        Val += row.Mar
                    End If
                    If Period.Month = 4 Then
                        Val += row.Apr
                    End If
                    If Period.Month = 5 Then
                        Val += row.Mei
                    End If
                    If Period.Month = 6 Then
                        Val += row.Jun
                    End If
                    If Period.Month = 7 Then
                        Val += row.Jul
                    End If
                    If Period.Month = 8 Then
                        Val += row.Agt
                    End If
                    If Period.Month = 9 Then
                        Val += row.Sept
                    End If
                    If Period.Month = 10 Then
                        Val += row.Okt
                    End If
                    If Period.Month = 11 Then
                        Val += row.Nov
                    End If
                    If Period.Month = 12 Then
                        Val += row.Des
                    End If
                    SN = row.Saldo_Normal
                    If SN = ClsMasterAkun.EnumSaldoNormal.Debet Then
                        .Debet = Val
                    Else
                        .Credit = Val
                    End If
                End With
                TBL.Rows.Add(rowBudget)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBLItem.Dispose()
            ObjCD = Nothing
        End Try
    End Sub

    Public Sub GetBudget_PL(ByVal PeriodStart As Date, ByVal PeriodEnd As Date, ByVal TBL As DseLaporanTrialBalance.Trial_BalanceDataTable)
        Dim vPrimaryBudget As Double = -1
        Dim DTA As New DseMasterBudgetTableAdapters.DataTable1TableAdapter
        Dim TBLItem As New DseMasterBudget.DataTable1DataTable
        Dim row As DseMasterBudget.DataTable1Row
        Dim rowBudget As DseLaporanTrialBalance.Trial_BalanceRow
        Dim Val As Double = 0
        Dim SN As ClsMasterAkun.EnumSaldoNormal = ClsMasterAkun.EnumSaldoNormal.Debet
        Dim ObjCD As New Core_Module.ClsCekData
        DTA.ClearBeforeFill = False
        Dim DummyPeriod As Date
        Dim StartMonth, EndMonth As Integer 

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBLItem.Rows.Clear()
            TBL.Rows.Clear()

            DummyPeriod = ObjCD.GetFirstDateofTheMonth(PeriodStart)

            While ObjCD.GetFisrtDateOfTheYear(DummyPeriod) <= ObjCD.GetFisrtDateOfTheYear(PeriodEnd)
                vPrimaryBudget = Me.GetPrimary(ObjCD.GetFisrtDateOfTheYear(DummyPeriod), FrmMasterBudget.EnumFormatYearly.Profit_N_Loss)
                TBLItem.Rows.Clear()
                DTA.Fill(TBLItem, vPrimaryBudget)
                If DummyPeriod.Year = PeriodEnd.Year Then
                    StartMonth = DummyPeriod.Month
                    EndMonth = PeriodEnd.Month
                Else
                    StartMonth = DummyPeriod.Month
                    EndMonth = 12
                End If

                For Each row In TBLItem.Rows
                    TBL.DefaultView.RowFilter = "[Primary Akun] = " + row.Primary_Akun.ToString
                    Val = 0

                    If TBL.DefaultView.Count = 0 Then
                        rowBudget = TBL.NewTrial_BalanceRow
                        With rowBudget
                            .Primary_Akun = row.Primary_Akun
                            .Saldo_Normal = row.Saldo_Normal
                            .Debet = 0
                            .Credit = 0
                            .Akun = ""
                            .Header = row.Header
                            .Level_Perkiraan = row.Level_Perkiraan
                            .Beginning_Balance = row.Beginning_Balance
                            .Saldo = 0
                            .Aktif = row.Aktif
                            .Description = row.Description
                            .Tipe_Perkiraan = row.Tipe_Perkiraan
                            .Urutan = row.Urutan
                            .Account_Code = row.Account_Code
                        End With
                        TBL.Rows.Add(rowBudget)
                    Else
                        rowBudget = TBL.DefaultView(0).Row
                    End If

                    SN = row.Saldo_Normal
                    Val = 0
                    If StartMonth <= 1 And 1 <= EndMonth Then
                        Val += row.Januari
                    End If
                    If StartMonth <= 2 And 2 <= EndMonth Then
                        Val += row.Feb
                    End If
                    If StartMonth <= 3 And 3 <= EndMonth Then
                        Val += row.Mar
                    End If
                    If StartMonth <= 4 And 4 <= EndMonth Then
                        Val += row.Apr
                    End If
                    If StartMonth <= 5 And 5 <= EndMonth Then
                        Val += row.Mei
                    End If
                    If StartMonth <= 6 And 6 <= EndMonth Then
                        Val += row.Jun
                    End If
                    If StartMonth <= 7 And 7 <= EndMonth Then
                        Val += row.Jul
                    End If
                    If StartMonth <= 8 And 8 <= EndMonth Then
                        Val += row.Agt
                    End If
                    If StartMonth <= 9 And 9 <= EndMonth Then
                        Val += row.Sept
                    End If
                    If StartMonth <= 10 And 10 <= EndMonth Then
                        Val += row.Okt
                    End If
                    If StartMonth <= 11 And 11 <= EndMonth Then
                        Val += row.Nov
                    End If
                    If StartMonth <= 12 And 12 <= EndMonth Then
                        Val += row.Des
                    End If
                    If SN = ClsMasterAkun.EnumSaldoNormal.Debet Then
                        rowBudget.Debet += Val
                    Else
                        rowBudget.Credit += Val
                    End If
                Next

                DummyPeriod = ObjCD.GetLastDateOfTheYear(DummyPeriod) + TimeSpan.FromDays(1)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBLItem.Dispose()
            ObjCD = Nothing
        End Try
    End Sub

    Public Function GetPrimary(ByVal vFinancialYear As Date, ByVal vTipeBudget As FrmMasterBudget.EnumFormatYearly) As Double
        Dim DTA As New DseMasterBudgetTableAdapters.tblmaster_budgetTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vFinancialYear, vTipeBudget).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary(vFinancialYear, vTipeBudget).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vDGV As DataGridView, ByVal ij As FrmMasterBudget, _
    ByVal vPrimaryBudget As Double, ByVal Format As FrmMasterBudget.EnumFormatYearly) As Boolean
        Dim DTAMasterAcc As New DseMasterBudgetTableAdapters.Master_AccTableAdapter
        Dim DTAItem As New DseMasterBudgetTableAdapters.tblmaster_budget_itemTableAdapter
        Dim TBL_MasterAkun As New DseMasterBudget.Master_AccDataTable
        Dim TBL_Item As New DseMasterBudget.tblmaster_budget_itemDataTable
        Dim row_Akun As DseMasterBudget.Master_AccRow
        Dim row_Item As DseMasterBudget.tblmaster_budget_itemRow
        Dim filter As String = ""
        Dim row_DGV As DataGridViewRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim TBL_Header As New DseMasterBudget.tblmaster_budgetDataTable
        Dim row_Header As DseMasterBudget.tblmaster_budgetRow
        Dim DTA_Budget As New DseMasterBudgetTableAdapters.tblmaster_budgetTableAdapter
        Dim Def_Prec As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            DTAMasterAcc.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Budget.Connection.ConnectionString = Core_Module.GetConstring

            Def_Prec = ObjCur.GetDecimal_DefaultCurrency
            'cari data header
            DTA_Budget.FillBy_Primary(TBL_Header, vPrimaryBudget)
            If TBL_Header.Rows.Count > 0 Then
                row_Header = TBL_Header.Rows(0)
                With row_Header
                    Me.mPrimary = .Primary
                    Me.mTimestamp = .Timestamp
                    Me.mFinancialYear = .Period
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                End With
            Else
                Me.mPrimary = -1
                Me.mTimestamp = Now
                Me.mFinancialYear = #1/1/1950#
                Me.mCreatedBy = ""
                Me.mCreatedDate = #1/1/1950#
                Me.mModifiedBy = ""
                Me.mModifiedDate = #1/1/1950#
            End If

            'kosongkan dgv
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            vDGV.Rows.Clear()

            'isikan data master akun + item
            If Format = FrmMasterBudget.EnumFormatYearly.Balance_Sheet Then
                DTAMasterAcc.FillBy_AkunBalanceSheet(TBL_MasterAkun)
            Else
                DTAMasterAcc.FillBy_AkunP_L(TBL_MasterAkun)
            End If
            DTAItem.FillBy_PrimaryBudget(TBL_Item, vPrimaryBudget)

            'gabungkan data master akun dan budget, transfer ke dalam dgv
            For Each row_Akun In TBL_MasterAkun.Rows
                filter = "[Primary Akun] =" + row_Akun.Primary_Akun.ToString
                TBL_Item.DefaultView.RowFilter = filter

                vDGV.Rows.Add()
                row_DGV = vDGV.Rows(vDGV.Rows.Count - 1)

                If TBL_Item.DefaultView.Count = 0 Then
                    'tambahkan data kosong
                    With row_DGV
                        .Cells(ij.ColAccount.Index).Value = row_Akun.Akun
                        .Cells(ij.ColHeader.Index).Value = row_Akun.Header
                        .Cells(ij.ColLevel.Index).Value = row_Akun.Level_Perkiraan
                        .Cells(ij.ColPrimaryAkun.Index).Value = row_Akun.Primary_Akun
                        .Cells(ij.ColRemarks.Index).Value = ""
                        .Cells(ij.ColJan.Index).Value = ij.ObjCD.DigitGrouping(0, Def_Prec)
                        .Cells(ij.ColFeb.Index).Value = ij.ObjCD.DigitGrouping(0, Def_Prec)
                        .Cells(ij.ColMar.Index).Value = ij.ObjCD.DigitGrouping(0, Def_Prec)
                        .Cells(ij.ColApr.Index).Value = ij.ObjCD.DigitGrouping(0, Def_Prec)
                        .Cells(ij.ColMay.Index).Value = ij.ObjCD.DigitGrouping(0, Def_Prec)
                        .Cells(ij.ColJune.Index).Value = ij.ObjCD.DigitGrouping(0, Def_Prec)
                        .Cells(ij.ColJul.Index).Value = ij.ObjCD.DigitGrouping(0, Def_Prec)
                        .Cells(ij.ColAgt.Index).Value = ij.ObjCD.DigitGrouping(0, Def_Prec)
                        .Cells(ij.ColSept.Index).Value = ij.ObjCD.DigitGrouping(0, Def_Prec)
                        .Cells(ij.ColOct.Index).Value = ij.ObjCD.DigitGrouping(0, Def_Prec)
                        .Cells(ij.ColNov.Index).Value = ij.ObjCD.DigitGrouping(0, Def_Prec)
                        .Cells(ij.ColDec.Index).Value = ij.ObjCD.DigitGrouping(0, Def_Prec)
                    End With
                Else
                    'tambahkan data sesuai dengan budget
                    row_Item = TBL_Item.DefaultView(0).Row
                    With row_DGV
                        .Cells(ij.ColAccount.Index).Value = row_Akun.Akun
                        .Cells(ij.ColHeader.Index).Value = row_Akun.Header
                        .Cells(ij.ColLevel.Index).Value = row_Akun.Level_Perkiraan
                        .Cells(ij.ColPrimaryAkun.Index).Value = row_Akun.Primary_Akun
                        .Cells(ij.ColRemarks.Index).Value = ""
                        .Cells(ij.ColJan.Index).Value = ij.ObjCD.DigitGrouping(row_Item.Januari, Def_Prec)
                        .Cells(ij.ColFeb.Index).Value = ij.ObjCD.DigitGrouping(row_Item.Feb, Def_Prec)
                        .Cells(ij.ColMar.Index).Value = ij.ObjCD.DigitGrouping(row_Item.Mar, Def_Prec)
                        .Cells(ij.ColApr.Index).Value = ij.ObjCD.DigitGrouping(row_Item.Apr, Def_Prec)
                        .Cells(ij.ColMay.Index).Value = ij.ObjCD.DigitGrouping(row_Item.Mei, Def_Prec)
                        .Cells(ij.ColJune.Index).Value = ij.ObjCD.DigitGrouping(row_Item.Jun, Def_Prec)
                        .Cells(ij.ColJul.Index).Value = ij.ObjCD.DigitGrouping(row_Item.Jul, Def_Prec)
                        .Cells(ij.ColAgt.Index).Value = ij.ObjCD.DigitGrouping(row_Item.Agt, Def_Prec)
                        .Cells(ij.ColSept.Index).Value = ij.ObjCD.DigitGrouping(row_Item.Sept, Def_Prec)
                        .Cells(ij.ColOct.Index).Value = ij.ObjCD.DigitGrouping(row_Item.Okt, Def_Prec)
                        .Cells(ij.ColNov.Index).Value = ij.ObjCD.DigitGrouping(row_Item.Nov, Def_Prec)
                        .Cells(ij.ColDec.Index).Value = ij.ObjCD.DigitGrouping(row_Item.Des, Def_Prec)
                    End With
                End If
            Next

            'sampe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAMasterAcc.Dispose()
            DTAItem.Dispose()
            TBL_Item.Dispose()
            TBL_MasterAkun.Dispose()
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            TBL_Header.Dispose()
            DTA_Budget.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
            ObjCur = Nothing
        End Try
    End Function

    Public Function GetJml_ByPrimaryAkun(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterBudgetTableAdapters.tblmaster_budget_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlByPrimaryAkun(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByPrimaryAkun(vPrimaryAkun).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiDGV_Tahunan(ByVal vDGV As DataGridView, ByVal ij As FrmMasterBudget, ByVal vFinancialYear As Date, ByVal vTipeBudget As FrmMasterBudget.EnumFormatYearly)
        Dim DTAMasterAcc As New DseMasterBudgetTableAdapters.Master_AccTableAdapter
        Dim DTAItem As New DseMasterBudgetTableAdapters.tblmaster_budget_item1TableAdapter
        Dim TBL_MasterAkun As New DseMasterBudget.Master_AccDataTable
        Dim TBL_Item As New DseMasterBudget.tblmaster_budget_item1DataTable
        Dim row_Akun As DseMasterBudget.Master_AccRow
        Dim row_Item As DseMasterBudget.tblmaster_budget_item1Row
        Dim filter As String = ""
        Dim row_DGV As DataGridViewRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim Dummy As Date = #1/1/1950#
        Dim Def_Prec As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim TBL As New DseLaporanTrialBalance.Trial_BalanceDataTable

        Try
            DTAMasterAcc.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            Def_Prec = ObjCur.GetDecimal_DefaultCurrency

            'kosongkan dgv
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            vDGV.Rows.Clear()

            'isikan data master akun + item
            If vTipeBudget = FrmMasterBudget.EnumFormatYearly.Balance_Sheet Then
                DTAMasterAcc.FillBy_AkunBalanceSheet(TBL_MasterAkun)
            Else
                DTAMasterAcc.FillBy_AkunP_L(TBL_MasterAkun)
            End If
            DTAItem.Fill(TBL_Item, Me.GetPrimary(vFinancialYear, Me.mTipeBudget))

            'gabungkan data master akun dan budget, transfer ke dalam dgv, isikan data tahun ybs
            For Each row_Akun In TBL_MasterAkun.Rows
                filter = "[Primary Akun] =" + row_Akun.Primary_Akun.ToString
                TBL_Item.DefaultView.RowFilter = filter

                vDGV.Rows.Add()
                row_DGV = vDGV.Rows(vDGV.Rows.Count - 1)

                If TBL_Item.DefaultView.Count = 0 Then
                    'tambahkan data kosong
                    With row_DGV
                        .Cells(ij.ColAccount_Yearly.Index).Value = row_Akun.Akun
                        .Cells(ij.ColPrimaryAkunYearly.Index).Value = row_Akun.Primary_Akun
                        .Cells(ij.ColCurYear_Budget.Index).Value = ij.ObjCD.DigitGrouping(0, Def_Prec)
                        .Cells(ij.ColCurrenctYear_Rea.Index).Value = ij.ObjCD.DigitGrouping(0, Def_Prec)
                    End With
                Else
                    'tambahkan data sesuai dengan budget
                    row_Item = TBL_Item.DefaultView(0).Row
                    With row_DGV
                        .Cells(ij.ColAccount_Yearly.Index).Value = row_Akun.Akun
                        .Cells(ij.ColPrimaryAkunYearly.Index).Value = row_Akun.Primary_Akun
                        .Cells(ij.ColCurYear_Budget.Index).Value = ij.ObjCD.DigitGrouping(row_Item.Total, Def_Prec)
                        .Cells(ij.ColCurrenctYear_Rea.Index).Value = ij.ObjCD.DigitGrouping(0, Def_Prec)
                    End With
                End If
            Next
            ij.ColCurYear_Budget.HeaderText = "Budget " + Format(vFinancialYear, "yyyy")

            'isikan budget 1 tahun sebelumnya
            Dummy = ij.ObjCD.GetFisrtDateOfTheYear(vFinancialYear - TimeSpan.FromDays(1))
            Me.IsiBudget(Dummy, ij.ColYear1_Bud.Index, DTAItem, TBL_Item, _
            vDGV, ij.ObjCD, ij.ColPrimaryAkunYearly.Index, vTipeBudget)

            'isikan budget 2 tahun sebelumnya
            Dummy = ij.ObjCD.GetFisrtDateOfTheYear(vFinancialYear - TimeSpan.FromDays(1 + 366))
            Me.IsiBudget(Dummy, ij.ColYear2_Budget.Index, DTAItem, TBL_Item, _
            vDGV, ij.ObjCD, ij.ColPrimaryAkunYearly.Index, vTipeBudget)

            'isikan budget 3 tahun sebelumnya
            Dummy = ij.ObjCD.GetFisrtDateOfTheYear(vFinancialYear - TimeSpan.FromDays(1 + 366 + 366))
            Me.IsiBudget(Dummy, ij.ColYear3_Bud.Index, DTAItem, TBL_Item, _
            vDGV, ij.ObjCD, ij.ColPrimaryAkunYearly.Index, vTipeBudget)

            'isikan budget 1 tahun setelahnya
            Dummy = ij.ObjCD.GetFisrtDateOfTheYear(vFinancialYear + TimeSpan.FromDays(1 + 366))
            Me.IsiBudget(Dummy, ij.ColNextYear_Bud.Index, DTAItem, TBL_Item, _
            vDGV, ij.ObjCD, ij.ColPrimaryAkunYearly.Index, vTipeBudget)

            'isikan realisasi
            If vTipeBudget = FrmMasterBudget.EnumFormatYearly.Balance_Sheet Then
                Me.IsiRealisasiBS(vFinancialYear, TBL, vDGV, ij.ColCurrenctYear_Rea.Index, ij.ColPrimaryAkunYearly.Index, Def_Prec)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAMasterAcc.Dispose()
            DTAItem.Dispose()
            TBL_Item.Dispose()
            TBL_MasterAkun.Dispose()
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Windows.Forms.Cursor.Current = Cursors.Arrow
            ObjCur = Nothing
            TBL = Nothing
        End Try
    End Sub

    Private Sub IsiRealisasiBS(ByVal vEndDateOfYear As Date, _
    ByVal TBL As DseLaporanTrialBalance.Trial_BalanceDataTable, ByVal vDGV As DataGridView, ByVal vIndexRealisasi As Integer, _
    ByVal vIndexPrimary As Integer, ByVal vDef_Prec As Integer)
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowBeginningBalance As DseLaporanTrialBalance.tbl_beginning_balanceRow
        Dim filter As String = ""
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim rowDGV As DataGridViewRow
        vEndDateOfYear = ObjCD.GetLasDateOfTheMonth(vEndDateOfYear)

        Try
            Me.DTA_Beginning_Balance.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTA_TrialBalance.Connection.ConnectionString = Core_Module.GetConstring

            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Cari semua akun yang ada dengan begining balance pada awal period (kalo start period Jul 2010 --> begining balance
            '   adalah saldo buku besar pada 30 jun 2010
            Me.DTA_Beginning_Balance.Fill(Me.TBL_Beginning_Balance, vEndDateOfYear)

            '2. Pindahkan ke tabel trial balance
            Me.DTA_TrialBalance.Fill(TBL)

            '2.b. Buang akun selain asset, liabilities, equity
            j = 0
            While j < TBL.Rows.Count
                row = TBL.Rows(j)
                If row.Tipe_Perkiraan > 2 Then
                    TBL.Rows.Remove(row)
                Else
                    j += 1
                End If
            End While

            For Each row In TBL.Rows
                filter = "Primary =" + row.Primary_Akun.ToString
                Me.TBL_Beginning_Balance.DefaultView.RowFilter = filter
                If Me.TBL_Beginning_Balance.DefaultView.Count > 0 Then
                    rowBeginningBalance = Me.TBL_Beginning_Balance.DefaultView(0).Row
                    row.Debet += rowBeginningBalance.Debet
                    row.Credit += rowBeginningBalance.Kredit

                    If row.Saldo_Normal = 0 Then
                        row.Debet = row.Beginning_Balance + row.Debet - row.Credit
                        row.Credit = 0
                    Else
                        row.Credit = row.Beginning_Balance + row.Credit - row.Debet
                        row.Debet = 0
                    End If
                End If
            Next

            '3. Hitung sum untuk setiap data header 
            Me.UpdateSaldoHeader(TBL)

            For Each rowDGV In vDGV.Rows
                filter = "[Primary Akun] =" + rowDGV.Cells(vIndexPrimary).Value.ToString
                TBL.DefaultView.RowFilter = filter
                row = TBL.DefaultView(0).Row

                If TBL.DefaultView.Count > 0 Then
                    If row.Saldo_Normal = 0 Then
                        rowDGV.Cells(vIndexRealisasi).Value = Me.ObjCD.DigitGrouping(row.Debet, vDef_Prec)
                    Else
                        rowDGV.Cells(vIndexRealisasi).Value = Me.ObjCD.DigitGrouping(row.Credit, vDef_Prec)
                    End If
                Else
                    rowDGV.Cells(vIndexRealisasi).Value = Me.ObjCD.DigitGrouping(0, vDef_Prec)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateSaldoHeader(ByVal vTBL As DseLaporanTrialBalance.Trial_BalanceDataTable)
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim i As Integer = 0

        Try
            For Each row In vTBL.Rows
                If row.Header = 1 Then
                    If row.Saldo_Normal = 0 Then
                        row.Debet = Me.GetSaldoHeader(i, row.Saldo_Normal, row.Level_Perkiraan, vTBL)
                    Else
                        row.Credit = Me.GetSaldoHeader(i, row.Saldo_Normal, row.Level_Perkiraan, vTBL)
                    End If
                End If

                i += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetSaldoHeader(ByVal vRowIndex As Double, ByVal vDebetKredit As Integer, ByVal vLevel As Integer, _
  ByVal vTBL As DseLaporanTrialBalance.Trial_BalanceDataTable) As Double
        Dim TargetIndex, CurrentIndex As Integer
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim RetVal As Double = 0

        Try
            CurrentIndex = vRowIndex + 1
            While CurrentIndex <= vTBL.Rows.Count - 1
                row = vTBL.Rows(CurrentIndex)
                If row.Level_Perkiraan <= vLevel Then
                    TargetIndex = CurrentIndex - 1
                    Exit While
                ElseIf CurrentIndex = vTBL.Rows.Count - 1 Then
                    TargetIndex = CurrentIndex
                    Exit While
                End If
                CurrentIndex += 1
            End While

            For i As Integer = vRowIndex + 1 To TargetIndex
                row = vTBL.Rows(i)
                If row.Header = 1 And row.Level_Perkiraan = vLevel + 1 Then
                    RetVal += Me.GetSaldoHeader(i, vDebetKredit, row.Level_Perkiraan, vTBL)
                ElseIf row.Level_Perkiraan = vLevel + 1 Then
                    If vDebetKredit = 0 Then
                        RetVal += CDec(row.Debet)
                    Else
                        RetVal += CDec(row.Credit)
                    End If
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsiBudget(ByVal vFinancialYear As Date, ByVal vColIndex As Integer, _
    ByVal DTAItem As DseMasterBudgetTableAdapters.tblmaster_budget_item1TableAdapter, _
    ByVal TBL_Item As DseMasterBudget.tblmaster_budget_item1DataTable, _
    ByVal vDGV As DataGridView, ByVal ObjCD As Core_Module.ClsCekData, ByVal vColPrimaryAkun As Integer, ByVal vTipeBudget As EnumTipeBudget)
        Dim Filter As String = ""
        Dim row_DGV As DataGridViewRow
        Dim row_Item As DseMasterBudget.tblmaster_budget_item1Row
        Dim Def_Prec As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            Def_Prec = ObjCur.GetDecimal_DefaultCurrency
            DTAItem.Fill(TBL_Item, Me.GetPrimary(vFinancialYear, vTipeBudget))
            For Each row_DGV In vDGV.Rows
                Filter = "[Primary Akun] =" + row_DGV.Cells(vColPrimaryAkun).Value.ToString
                TBL_Item.DefaultView.RowFilter = Filter
                If TBL_Item.DefaultView.Count = 0 Then
                    'set kosong
                    With row_DGV
                        .Cells(vColIndex).Value = ObjCD.DigitGrouping(0, Def_Prec)
                    End With
                Else
                    'modifikasi data sesuai dengan budget
                    row_Item = TBL_Item.DefaultView(0).Row
                    With row_DGV
                        .Cells(vColIndex).Value = ObjCD.DigitGrouping(row_Item.Total, Def_Prec)
                    End With
                End If
            Next
            vDGV.Columns(vColIndex).HeaderText = "Budget " + Format(vFinancialYear, "yyyy")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCur = Nothing
        End Try
    End Sub
End Class


