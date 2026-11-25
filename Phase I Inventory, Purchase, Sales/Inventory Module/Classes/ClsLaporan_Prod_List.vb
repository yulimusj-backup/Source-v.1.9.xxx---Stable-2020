Option Explicit On
Option Strict Off
Public Class ClsLaporan_Prod_List
    'membernya
    Private mPrimary As Double = -1
    Private mDesc As String = ""
    Private mFormat1 As EnumFormat
    Private mFormat2 As EnumFormat2
    Private mGroupBy As EnumGroupBy
    Private mFromDate As Date = Now
    Private mToDate As Date = Now
    Private mGroupLevel As Core_Module.ClsMasterProductAccountGroup.EnumGroupLevel
    Private mFilterByWorkUnit As Integer
    Private mFilterByProcess As Integer
    Private mFilterByProductGroup As Integer
    Private mFilterByWorkshift As Integer
    Private mShowsubTotal As Integer
    Private mTimestamp As Date

#Region "Enum disini"
    Public Enum EnumFormat As Integer
        Recap
        Recap_w_group
        Detail
    End Enum

    Public Enum EnumFormat2 As Integer
        Input
        Output
        Both
    End Enum

    Public Enum EnumGroupBy As Integer
        _Date
        Document_No
        Work_Unit
        Work_Shift
        Process
    End Enum

    Public Enum EnumFilter As Integer
        No
        YEs
    End Enum

    Public Enum EnumShowSubTotal As Integer
        No
        Yes
    End Enum

    Public Enum EnumDirection
        Ascending
        Descending
    End Enum
#End Region

#Region "Property disini"
    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property Desc() As String
        Get
            Return Me.mDesc
        End Get
        Set(ByVal value As String)
            mDesc = value
        End Set
    End Property

    Public Property Format1() As Integer
        Get
            Return Me.mFormat1
        End Get
        Set(ByVal value As Integer)
            mFormat1 = value
        End Set
    End Property

    Public Property Format2() As Integer
        Get
            Return Me.mFormat2
        End Get
        Set(ByVal value As Integer)
            mFormat2 = value
        End Set
    End Property

    Public Property GroupBy() As Integer
        Get
            Return Me.mGroupBy
        End Get
        Set(ByVal value As Integer)
            mGroupBy = value
        End Set
    End Property

    Public Property FromDate() As Date
        Get
            Return Me.mFromDate
        End Get
        Set(ByVal value As Date)
            mFromDate = value
        End Set
    End Property

    Public Property ToDate() As Date
        Get
            Return Me.mToDate
        End Get
        Set(ByVal value As Date)
            mToDate = value
        End Set
    End Property

    Public Property GroupLevel() As Integer
        Get
            Return Me.mGroupLevel
        End Get
        Set(ByVal value As Integer)
            mGroupLevel = value
        End Set
    End Property

    Public Property FilterByWorkUnit() As Integer
        Get
            Return Me.mFilterByWorkUnit
        End Get
        Set(ByVal value As Integer)
            mFilterByWorkUnit = value
        End Set
    End Property

    Public Property FilterByWorkshift() As Integer
        Get
            Return Me.mFilterByWorkshift
        End Get
        Set(ByVal value As Integer)
            mFilterByWorkshift = value
        End Set
    End Property

    Public Property FilterByProcess() As Integer
        Get
            Return Me.mFilterByProcess
        End Get
        Set(ByVal value As Integer)
            mFilterByProcess = value
        End Set
    End Property

    Public Property FilterByProductGroup() As Integer
        Get
            Return Me.mFilterByProductGroup
        End Get
        Set(ByVal value As Integer)
            mFilterByProductGroup = value
        End Set
    End Property

    Public Property ShowsubTotal() As Integer
        Get
            Return Me.mShowsubTotal
        End Get
        Set(ByVal value As Integer)
            mShowsubTotal = value
        End Set
    End Property

    Public Property Timestamp() As Date
        Get
            Return Me.mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
        End Set
    End Property
#End Region

#Region "Simpan hapus dan isi dgv Item disini"
    Public Function SimpanItemDGVSort(ByVal vPrimaryTransaksi As Double, _
                                ByVal vDGVItem As DataGridView, _
                                ByVal vFrm As FrmLaporanProd_List) As Boolean

        Dim DTAItem As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling_item_sortTableAdapter
        Dim row As DataGridViewRow
        Dim ij As Integer = ClsLaporan_Prod_List.EnumDirection.Ascending

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For Each row In vDGVItem.Rows
                If row.Cells(vFrm.ColSort_Direction.Index).Value = "Ascending" Then
                    ij = ClsLaporan_Prod_List.EnumDirection.Ascending
                Else
                    ij = ClsLaporan_Prod_List.EnumDirection.Descending
                End If
                DTAItem.Insert(vPrimaryTransaksi, CStr(row.Cells(vFrm.ColSort_Desc.Index).Value), CInt(row.Cells(vFrm.ColSort_Order.Index).Value), _
                   ij)
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function SimpanItemDGVProcess(ByVal vPrimaryTransaksi As Double, _
                                ByVal vDGVItem As DataGridView, _
                                ByVal vFrm As FrmLaporanProd_List) As Boolean

        Dim DTAItem As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling_item_filter_processTableAdapter
        Dim row As DataGridViewRow

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If vFrm.chkProcess.Checked = True Then
                vDGVItem.AllowUserToAddRows = False
                For Each row In vDGVItem.Rows
                    If row.Cells(vFrm.ColProcess_Code.Index).Value = "" Then Exit For
                    With row
                        DTAItem.Insert(vPrimaryTransaksi, .Cells(vFrm.ColProcess_PrimaryOriginal.Index).Value, .Cells(vFrm.ColProcess_Code.Index).Value, .Cells(vFrm.ColProcess_Desc.Index).Value)
                    End With
                Next
                vDGVItem.AllowUserToAddRows = True
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function SimpanItemDGVWorkUnit(ByVal vPrimaryTransaksi As Double, _
                                ByVal vDGVItem As DataGridView, _
                                ByVal vFrm As FrmLaporanProd_List) As Boolean

        Dim DTAItem As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling_item_filter_workunitTableAdapter
        Dim row As DataGridViewRow

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If vFrm.chkWorkUnit.Checked = True Then
                vDGVItem.AllowUserToAddRows = False
                For Each row In vDGVItem.Rows
                    If row.Cells(vFrm.ColWorkUnit_Code.Index).Value = "" Then Exit For
                    With row
                        DTAItem.Insert(vPrimaryTransaksi, .Cells(vFrm.ColWorkUnit_PrimaryOriginal.Index).Value, .Cells(vFrm.ColWorkUnit_Code.Index).Value, .Cells(vFrm.ColWorkUnit_Desc.Index).Value)
                    End With
                Next
                vDGVItem.AllowUserToAddRows = True
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function SimpanItemDGVWorkshift(ByVal vPrimaryTransaksi As Double, _
                                ByVal vDGVItem As DataGridView, _
                                ByVal vFrm As FrmLaporanProd_List) As Boolean

        Dim DTAItem As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling_item_filter_workshiftTableAdapter
        Dim row As DataGridViewRow

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If vFrm.ChkWorkShift.Checked = True Then
                vDGVItem.AllowUserToAddRows = False
                For Each row In vDGVItem.Rows
                    If row.Cells(vFrm.ColWorkShift_Code.Index).Value = "" Then Exit For
                    With row
                        DTAItem.Insert(vPrimaryTransaksi, .Cells(vFrm.ColWorkShift_PrimaryOriginal.Index).Value, .Cells(vFrm.ColWorkShift_Code.Index).Value, .Cells(vFrm.ColWorkShift_Desc.Index).Value)
                    End With
                Next
                vDGVItem.AllowUserToAddRows = True
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function SimpanItemDGVProductGroup(ByVal vPrimaryTransaksi As Double, _
                                ByVal vDGVItem As DataGridView, _
                                ByVal vFrm As FrmLaporanProd_List) As Boolean

        Dim DTAItem As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling_item_filter_product_groupTableAdapter
        Dim row As DataGridViewRow

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If vFrm.ChkProductGroup.Checked = True Then
                vDGVItem.AllowUserToAddRows = False
                For Each row In vDGVItem.Rows
                    If row.Cells(vFrm.ColGroup_Code.Index).Value = "" Then Exit For
                    With row
                        DTAItem.Insert(vPrimaryTransaksi, .Cells(vFrm.ColGroup_PrimaryOriginal.Index).Value, .Cells(vFrm.ColGroup_Code.Index).Value, .Cells(vFrm.ColGroup_Desc.Index).Value)
                    End With
                Next
                vDGVItem.AllowUserToAddRows = True
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function HapusItemDGVSort(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTAItem As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling_item_sortTableAdapter
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Delete(vPrimaryTransaksi)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function HapusItemDGVProcess(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTAItem As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling_item_filter_processTableAdapter
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Delete(vPrimaryTransaksi)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function HapusItemDGVWorkUnit(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTAItem As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling_item_filter_workunitTableAdapter
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Delete(vPrimaryTransaksi)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function HapusItemDGVWorkshift(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTAItem As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling_item_filter_workshiftTableAdapter
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Delete(vPrimaryTransaksi)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function HapusItemDGVProductGroup(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTAItem As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling_item_filter_product_groupTableAdapter
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Delete(vPrimaryTransaksi)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function IsiDGVSort(ByVal vFrm As FrmLaporanProd_List, ByVal vPrimaryTransaksi As Double, _
                                  ByVal vDGV As DataGridView) As Boolean

        Dim DTA As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling_item_sortTableAdapter
        Dim TBL As New DseLaporanProd_List.tbltrans_setting_report_stockbundling_item_sortDataTable
        Dim row As DseLaporanProd_List.tbltrans_setting_report_stockbundling_item_sortRow
        Dim rowDGV As DataGridViewRow
        Dim ij As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.Rows.Clear()
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL
                    If row.Direction = ClsLaporan_Prod_List.EnumDirection.Ascending Then
                        ij = "Ascending"
                    Else
                        ij = "Descending"
                    End If
                    vDGV.Rows.Add()
                    rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                    With rowDGV
                        .Cells(vFrm.ColSort_Primary.Index).Value = row.Primary
                        .Cells(vFrm.ColSort_Desc.Index).Value = row.Sort_By
                        .Cells(vFrm.ColSort_Order.Index).Value = row.Sort_Order
                        .Cells(vFrm.ColSort_Direction.Index).Value = ij
                    End With
                Next
            End If
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGVProcess(ByVal vFrm As FrmLaporanProd_List, ByVal vPrimaryTransaksi As Double, _
                                  ByVal vDGV As DataGridView) As Boolean

        Dim DTA As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling_item_filter_processTableAdapter
        Dim TBL As New DseLaporanProd_List.tbltrans_setting_report_stockbundling_item_filter_processDataTable
        Dim row As DseLaporanProd_List.tbltrans_setting_report_stockbundling_item_filter_processRow
        Dim rowDGV As DataGridViewRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.Rows.Clear()
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                vDGV.AllowUserToAddRows = False

                For Each row In TBL
                    vDGV.Rows.Add()
                    rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                    With rowDGV
                        .Cells(vFrm.ColProcess_Primary.Index).Value = row.Primary
                        .Cells(vFrm.ColProcess_PrimaryOriginal.Index).Value = row.Primary_Original
                        .Cells(vFrm.ColProcess_Code.Index).Value = row.Code
                        .Cells(vFrm.ColProcess_Desc.Index).Value = row.Desc
                    End With
                Next
                vDGV.AllowUserToAddRows = True

            End If
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function IsiDGVWorkUnit(ByVal vFrm As FrmLaporanProd_List, ByVal vPrimaryTransaksi As Double, _
                                  ByVal vDGV As DataGridView) As Boolean

        Dim DTA As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling_item_filter_workunitTableAdapter
        Dim TBL As New DseLaporanProd_List.tbltrans_setting_report_stockbundling_item_filter_workunitDataTable
        Dim row As DseLaporanProd_List.tbltrans_setting_report_stockbundling_item_filter_workunitRow
        Dim rowDGV As DataGridViewRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.Rows.Clear()
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                vDGV.AllowUserToAddRows = False
                For Each row In TBL
                    vDGV.Rows.Add()
                    rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                    With rowDGV
                        .Cells(vFrm.ColWorkUnit_Primary.Index).Value = row.Primary
                        .Cells(vFrm.ColWorkUnit_PrimaryOriginal.Index).Value = row.Primary_Original
                        .Cells(vFrm.ColWorkUnit_Code.Index).Value = row.Code
                        .Cells(vFrm.ColWorkUnit_Desc.Index).Value = row.Desc
                    End With
                Next
                vDGV.AllowUserToAddRows = True
            End If
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function IsiDGVWorkshift(ByVal vFrm As FrmLaporanProd_List, ByVal vPrimaryTransaksi As Double, _
                                  ByVal vDGV As DataGridView) As Boolean

        Dim DTA As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling_item_filter_workshiftTableAdapter
        Dim TBL As New DseLaporanProd_List.tbltrans_setting_report_stockbundling_item_filter_workshiftDataTable
        Dim row As DseLaporanProd_List.tbltrans_setting_report_stockbundling_item_filter_workshiftRow
        Dim rowDGV As DataGridViewRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.Rows.Clear()
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                vDGV.AllowUserToAddRows = False

                For Each row In TBL
                    vDGV.Rows.Add()
                    rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                    With rowDGV
                        .Cells(vFrm.ColWorkShift_Primary.Index).Value = row.Primary
                        .Cells(vFrm.ColWorkShift_PrimaryOriginal.Index).Value = row.Primary_Original
                        .Cells(vFrm.ColWorkShift_Code.Index).Value = row.Code
                        .Cells(vFrm.ColWorkShift_Desc.Index).Value = row.Desc
                    End With
                Next
                vDGV.AllowUserToAddRows = True

            End If
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function IsiDGVProductGroup(ByVal vFrm As FrmLaporanProd_List, ByVal vPrimaryTransaksi As Double, _
                                  ByVal vDGV As DataGridView) As Boolean

        Dim DTA As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling_item_filter_product_groupTableAdapter
        Dim TBL As New DseLaporanProd_List.tbltrans_setting_report_stockbundling_item_filter_product_groupDataTable
        Dim row As DseLaporanProd_List.tbltrans_setting_report_stockbundling_item_filter_product_groupRow
        Dim rowDGV As DataGridViewRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.Rows.Clear()
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                vDGV.AllowUserToAddRows = False

                For Each row In TBL
                    vDGV.Rows.Add()
                    rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                    With rowDGV
                        .Cells(vFrm.ColGroup_Primary.Index).Value = row.Primary
                        .Cells(vFrm.ColGroup_PrimaryOriginal.Index).Value = row.Primary_Original
                        .Cells(vFrm.ColGroup_Code.Index).Value = row.Code
                        .Cells(vFrm.ColGroup_Desc.Index).Value = row.Desc
                    End With
                Next
                vDGV.AllowUserToAddRows = True
            End If
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

#End Region

    'Public Function GetNoBaru(ByVal vTanggalTransaksi As Date, ByVal vUserCode As String) As String
    'Dim DTA As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundlingTableAdapter
    'Dim TanggalAwal, TanggalAkhir As Date
    'Dim NoUrut As String
    'Dim ObjCD As New Core_Module.ClsCekData
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring

    '        TanggalAwal = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)
    '        TanggalAkhir = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
    '        If DTA.ScalarQuery_GetMax(TanggalAwal, TanggalAkhir).HasValue Then
    '            NoUrut = CStr(DTA.ScalarQuery_GetMax(TanggalAwal, TanggalAkhir).Value + 1)
    '            If NoUrut <= 0 Then
    '                NoUrut = 1
    '            End If
    '        Else
    '            NoUrut = 1
    '        End If
    '        If IsNumeric(NoUrut) Then
    '            Me.mUrutan = CDbl(NoUrut)
    '        End If
    '        While NoUrut.Length < 6
    '            NoUrut = "0" + NoUrut
    '        End While
    '        Return "Report - " + NoUrut '+ "/" + CStr(Format(vTanggalTransaksi, "MMM").ToUpper) + "/" + CStr(Format(vTanggalTransaksi, "yyyy"))
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return ""
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Function Simpan(ByVal vDGVSort As DataGridView, _
    ByVal vDGVProductGroup As DataGridView, _
    ByVal vDGVWorkUnit As DataGridView, _
    ByVal vDGVWorkshift As DataGridView, _
    ByVal vDGVProcess As DataGridView, _
    ByVal vfrm As FrmLaporanProd_List) As Boolean
        Dim DTA As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundlingTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.mDesc, Me.FromDate, Me.mToDate, Me.Format1, Me.Format2, Me.GroupBy, Me.GroupLevel, Me.ShowsubTotal, Me.FilterByProcess, _
            Me.FilterByWorkUnit, Me.FilterByProductGroup, Me.FilterByWorkshift, Now) > 0 Then
                Me.mPrimary = Me.GetPrimary(Me.mDesc)
                Me.SimpanItemDGVProcess(mPrimary, vDGVProcess, vfrm)
                Me.SimpanItemDGVWorkUnit(mPrimary, vDGVWorkUnit, vfrm)
                Me.SimpanItemDGVWorkshift(mPrimary, vDGVWorkshift, vfrm)
                Me.SimpanItemDGVProductGroup(mPrimary, vDGVProductGroup, vfrm)
                Me.SimpanItemDGVSort(mPrimary, vDGVSort, vfrm)
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vDGVSort As DataGridView, _
   ByVal vDGVProductGroup As DataGridView, _
   ByVal vDGVWorkUnit As DataGridView, _
   ByVal vDGVWorkshift As DataGridView, _
   ByVal vDGVProcess As DataGridView, _
   ByVal vfrm As FrmLaporanProd_List) As Boolean
        Dim DTA As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundlingTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Update(Me.mDesc, Me.FromDate, Me.mToDate, Me.Format1, Me.Format2, Me.GroupBy, Me.GroupLevel, Me.ShowsubTotal, Me.FilterByProcess, _
            Me.FilterByWorkUnit, Me.FilterByProductGroup, Me.FilterByWorkshift, Now, mPrimary) > 0 Then
                'Me.mPrimary = Me.GetPrimary(Me.mCode)
                Me.HapusItemDGVSort(mPrimary)
                Me.SimpanItemDGVSort(mPrimary, vDGVSort, vfrm)

                Me.HapusItemDGVProcess(mPrimary)
                Me.SimpanItemDGVProcess(mPrimary, vDGVProcess, vfrm)

                Me.HapusItemDGVProductGroup(mPrimary)
                Me.SimpanItemDGVProductGroup(mPrimary, vDGVProductGroup, vfrm)

                Me.HapusItemDGVWorkshift(mPrimary)
                Me.SimpanItemDGVWorkshift(mPrimary, vDGVWorkshift, vfrm)

                Me.HapusItemDGVWorkUnit(mPrimary)
                Me.SimpanItemDGVWorkUnit(mPrimary, vDGVWorkUnit, vfrm)
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiTBL(ByVal TBL As DseLaporanProd_List.tbltrans_setting_report_stockbundling1DataTable) As Boolean
        Dim DTA As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundling1TableAdapter
        Dim Row As DseLaporanProd_List.tbltrans_setting_report_stockbundling1Row
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.Fill(TBL)
            'isi dengan baris kosong
            Row = TBL.Newtbltrans_setting_report_stockbundling1Row
            With Row
                .Primary = -1
                .Report_Desc = ""
                .Timestamp = #1/1/1950#
            End With
            TBL.Addtbltrans_setting_report_stockbundling1Row(Row)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vDGVSort As DataGridView, _
    ByVal vDGVProductGroup As DataGridView, _
    ByVal vDGVWorkUnit As DataGridView, _
    ByVal vDGVWorkshift As DataGridView, _
    ByVal vDGVProcess As DataGridView, _
    ByVal vPrimary As Double, ByVal vFrm As FrmLaporanProd_List) As Boolean

        Dim DTA As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundlingTableAdapter
        Dim TBL As New DseLaporanProd_List.tbltrans_setting_report_stockbundlingDataTable
        Dim row As DseLaporanProd_List.tbltrans_setting_report_stockbundlingRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, CDec(vPrimary))
            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseLaporanProd_List.tbltrans_setting_report_stockbundlingRow)
                With row
                    Me.mPrimary = .Primary
                    Me.mDesc = .Report_Desc
                    Me.mFromDate = .From_Date
                    Me.mToDate = .To_Date
                    Me.mFormat1 = .Format1
                    Me.mFormat2 = .Format2
                    Me.mGroupBy = .Format_Group
                    Me.mGroupLevel = .Group_Level
                    Me.mShowsubTotal = .Show_sub_total
                    Me.mFilterByProcess = .Filter_Process
                    Me.mFilterByProductGroup = .Filter_Product_Group
                    Me.mFilterByWorkshift = .Filter_Workshift
                    Me.mFilterByWorkUnit = .Filter_Work_Unit
                    Me.mTimestamp = .Timestamp
                End With

                Me.IsiDGVSort(vFrm, vPrimary, vDGVSort)
                Me.IsiDGVProcess(vFrm, vPrimary, vDGVProcess)
                Me.IsiDGVProductGroup(vFrm, vPrimary, vDGVProductGroup)
                Me.IsiDGVWorkUnit(vFrm, vPrimary, vDGVWorkUnit)
                Me.IsiDGVWorkshift(vFrm, vPrimary, vDGVWorkshift)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundlingTableAdapter
        Dim Retval As Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Retval = DTA.ScalarQuery_GetPrimary(vCode)
            If Retval.HasValue = True Then
                Return Retval.Value
            Else
                Return -1
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahada(ByVal vDesc As String) As Boolean
        Dim DTA As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundlingTableAdapter
        Dim Retval As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Retval = DTA.ScalarQuery_GetJmlByDesc(vDesc)
            If Retval.HasValue = True Then
                If Retval.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
                Return False
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseLaporanProd_ListTableAdapters.tbltrans_setting_report_stockbundlingTableAdapter
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Retval = DTA.ScalarQuery_GetDesc(vPrimary)
            If String.IsNullOrEmpty(Retval) = False Then
                Return Retval
            Else
                Return ""
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
