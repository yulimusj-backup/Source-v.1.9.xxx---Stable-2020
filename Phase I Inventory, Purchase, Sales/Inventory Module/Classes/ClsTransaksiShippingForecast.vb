Option Explicit On
Option Strict Off

Public Class ClsTransaksiShippingForecast
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mNoUrut As Double = -1
    Private mNoTransaksi As String = ""
    Private mTanggalTransaksi As Date = Now
    Private mRemarks As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianMaster
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Public Enum EnumPrintAvailability As Integer
        All
        AvailableOnly
        NotAvailableOnly
    End Enum
    Public Enum EnumPrintGrouping As Integer
        Customer
        Product_Code
    End Enum
    Public Enum EnumOrderNo As Integer
        Sales_Order_Number
        Cust_PO_Number
        Both
    End Enum

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

    Public Function GetTanggalSF(ByVal vPrimarySF As Double) As Date
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecastTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetDate(vPrimarySF).HasValue = True Then
                Return DTA.ScalarQuery_GetDate(vPrimarySF).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
       ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
       Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecastTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByNoTransLIKE(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.ScalarQuery_GetJmlByTglDariSampai(vTanggalSampai, vTanggalDari).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                Return DTA.ScalarQuery_GetJmlByRemarksLIKE(Criteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoBaru1(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecastTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut As Double = 0
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cari tanggal sampai dan dari
            Me.ObjNumbering.GetTanggalDariSampai(TanggalStart, TanggalEnd, _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._SF, vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetUrutanMax(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetUrutanMax(TanggalEnd, TanggalStart).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            Return Me.ObjNumbering.GetNomorTransaksiBaru1(vTanggalTransaksi, NoUrut, _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._SF)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecastTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vNoTransaksi).Value > 0 Then
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

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
    ByVal vTBLItemStock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable, _
    ByVal vTBLItemPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable, _
    ByVal vTBLItemSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable) As Boolean
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecastTableAdapter
        Dim TBL As New DseTransaksiShippingForecast.tbltrans_shipping_forecastDataTable
        Dim row As DseTransaksiShippingForecast.tbltrans_shipping_forecastRow
        Dim ObjItem As New ClsTransaksiShippingForecast_Item

        Try
            'fil basis data
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mNoTransaksi = .document_no
                    Me.mNoUrut = .urutan
                    Me.mPrimary = .primary
                    Me.mRemarks = .remarks
                    Me.mTanggalTransaksi = ._date


                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With

                'salin data item 
                ObjItem.IsiDGV(Me.mPrimary, vTBLItemStock, vTBLItemPOB, vTBLItemSOB)

                'nilai balikan
                Return True
            Else
                Me.ResetObject(vTBLItemPOB, vTBLItemSOB, vTBLItemStock)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiShippingForecast.tbltrans_shipping_forecast1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecast1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_DocNo(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TglDariSampai(TBL, vTanggalSampai, vTanggalDari)
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_Remarks(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                Dim ij As New DataGridViewCellStyle
                ij.Format = "dd MMM yyyy"
                DGV.Columns(2).DefaultCellStyle = ij
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

    Public Overridable Function ApakahBolehEdit(ByVal vPrimarySF As Double, _
    ByVal vUserCode As String, Optional ByVal vStrictEdit As Boolean = True, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecastTableAdapter
        Dim CreatedBy As String = ""
        Dim ObjUser As New Core_Module.ClsUser

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            '0. jika strict edit maka cek dulu siapa yang membuat
            If vStrictEdit = True Then
                If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimarySF)) = True Then
                    CreatedBy = ""
                Else
                    CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimarySF)
                End If
                If CreatedBy = vUserCode Then
                Else
                    If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_MaterialPlanningDashboard_SF_OverrideEdit") = False Then
                        If TampilkanPesan = True Then
                            MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjUser = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimarySF As Double, ByVal vUserCode As String, _
    Optional ByVal vStrictEdit As Boolean = True, Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecastTableAdapter
        Dim ObjUser As New Core_Module.ClsUser

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'Suatu transaksi do tidak bisa diedit jika:
            '0. jika strict edit maka harus dicek siapa yang buat
            If vStrictEdit = True Then
                If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimarySF)) = True Then
                    CreatedBy = ""
                Else
                    CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimarySF)
                End If
                If CreatedBy = vUserCode Then
                Else
                    If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTrans_MaterialPlanningDashboard_SF_OverrideDelete") = False Then
                        If TampilkanPesan = True Then
                            MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjUser = Nothing
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean, _
    ByVal vDT As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable, _
    ByVal vPrimarySF As Double) As Boolean
        Try
            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._SF) = True Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        MsgBox("Document number duplication, please change document number.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Else
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._SF) = True Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        If Me.GetPrimary(Me.mNoTransaksi) <> Me.mPrimary Then
                            Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                        End If
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        If Me.GetPrimary(Me.mNoTransaksi) <> Me.mPrimary Then
                            MsgBox("Document number duplication, please change document number.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            End If

            '2. Jika tidak ada item
            If vDT.Rows.Count = 0 Then
                MsgBox("No item detected, process cancelled.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetPrimary(ByVal vNoSF As String) As Double
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecastTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vNoSF).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary(vNoSF).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal ij As FrmTrans_MaterialPlanningDashboard, _
    ByVal vDT_ItemPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable, _
    ByVal vDT_ItemSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable, _
    ByVal vDT_ItemStock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiSF_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vUserCode, ij, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemStock)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru1(ByVal vUserCode As String, ByVal ij As FrmTrans_MaterialPlanningDashboard, _
    ByVal vDT_ItemPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable, _
    ByVal vDT_ItemSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable, _
    ByVal vDT_ItemStock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable) As Boolean
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecastTableAdapter
        Dim ObjItem As New ClsTransaksiShippingForecast_Item
        Dim Urutan As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True, vDT_ItemSOB, Me.mPrimary) = True Then
                'insert query
                Urutan = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, _
                Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._SF, Me.mTanggalTransaksi)

                If DTA.Insert(Me.mNoTransaksi, Urutan, Me.mTanggalTransaksi, Me.mRemarks, _
                vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTransaksi).Value
                    Me.mNoUrut = Urutan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item baru
                    ObjItem.Simpan(Me.mPrimary, vDT_ItemStock, vDT_ItemPOB, vDT_ItemSOB, ij)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._SF) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    Return Me.SimpanBaru1(vUserCode, ij, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemStock)
                Else
                    MsgBox("Document Number duplication, please change document number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal ij As FrmTrans_MaterialPlanningDashboard, _
    ByVal vDT_ItemPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable, _
    ByVal vDT_ItemSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable, _
    ByVal vDT_ItemStock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiSF_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntrian.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vUserCode, ij, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemStock)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanPerubahan1(ByVal vUserCode As String, ByVal ij As FrmTrans_MaterialPlanningDashboard, _
    ByVal vDT_ItemPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable, _
    ByVal vDT_ItemSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable, _
    ByVal vDT_ItemStock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable) As Boolean
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecastTableAdapter
        Dim ObjItem As New ClsTransaksiShippingForecast_Item
        Dim Urutan As Double = -1
        Dim ObjEmail As New E_mail_Module.ClsTransEmail

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(False, vDT_ItemSOB, Me.mPrimary) = True Then
                Urutan = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, _
                Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._SF, Me.mTanggalTransaksi)

                If DTA.Update(Me.mNoTransaksi, Urutan, Me.mTanggalTransaksi, Me.mRemarks, vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mNoUrut = Urutan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.Simpan(Me.mPrimary, vDT_ItemStock, vDT_ItemPOB, vDT_ItemSOB, ij)

                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._SF) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    Return Me.SimpanPerubahan1(vUserCode, ij, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemStock)
                Else
                    MsgBox("Document number duplication, please change document number", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjEmail = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal UserCode As String, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiSF_Hapus, UserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > ObjAntri.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, UserCode, TampilkanPesan)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Public Overridable Function Hapus1(ByVal vPrimaryTransaksi As Double, ByVal UserCode As String, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecastTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimaryTransaksi, UserCode, TampilkanPesan) = True Then
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub ResetObject(ByVal vTBLPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable, _
    ByVal vTBLSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable, ByVal vTBLStock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable)
        Try
            Me.mCreatedBy = ""
            Me.mCreatedDate = Now
            Me.mModifiedBy = ""
            Me.mModifiedDate = Now
            Me.mNoTransaksi = ""
            Me.mNoUrut = -1
            Me.mPrimary = -1
            Me.mRemarks = ""
            Me.mTanggalTransaksi = Now
            Me.mTimestamp = Now

            vTBLPOB.Rows.Clear()
            vTBLSOB.Rows.Clear()
            vTBLStock.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetTimestamp(ByVal vPrimarySF As Double) As Date
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecastTableAdapter

        Try
            If DTA.ScalarQuery_GetTimestamp(vPrimarySF).HasValue = True Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimarySF).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDocumentNo(ByVal vPrimarySF As Double) As String
        Dim DTA As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecastTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDocNo(vPrimarySF)) = False Then
                Return DTA.ScalarQuery_GetDocNo(vPrimarySF)
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

    Public Function Cetak(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
  ByVal vprimarySF As Double, ByVal Grouping As EnumPrintGrouping, ByVal OrderNo As EnumOrderNo, ByVal vFilter As String, ByVal vLateOnly As Boolean, _
    ByVal vEnumPrintAv As EnumPrintAvailability, ByVal vAvailabilityDate As Date, ByVal vFilterAvDate As Boolean) As Boolean
        Dim report As New CrpTrans_MaterialPlanning_SF
        Dim DS As New DseTransaksiShippingForecast
        Dim DTAHeader As New DseTransaksiShippingForecastTableAdapters.tbltrans_shipping_forecast_itemdeliveryTableAdapter
        Dim TBLDummy As New DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable
        Dim rowDummy, rowCetak As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim DV As Data.DataRowView
        Dim Lanjut As Boolean = True

        Try
            DTAHeader.Connection.ConnectionString = Core_Module.GetConstring
            DTAHeader.FillBy_Primary(TBLDummy, vprimarySF)

            'filtering 
            TBLDummy.DefaultView.RowFilter = vFilter
            TBLDummy.DefaultView.Sort = "[customer name] ASC, [so date rec] ASC, [so number rec] ASC, [Product Code] ASC, [promise to deliver] ASC"
            For Each DV In TBLDummy.DefaultView
                Lanjut = True
                rowDummy = DV.Row
                If vLateOnly = True And Lanjut = True Then
                    If rowDummy.promise_to_deliver < rowDummy.estimate_delivery Then
                        'telat -> lanjut
                        Lanjut = True
                    Else
                        Lanjut = False
                    End If
                End If

                If Lanjut = True Then
                    If vEnumPrintAv = EnumPrintAvailability.AvailableOnly Then
                        'hanya yg available : qty available > qty sob dan tanggal available > dtp available
                        If rowDummy.qty_available >= rowDummy.qty_sob Then
                            If vFilterAvDate = False Then
                                Lanjut = True
                            Else
                                If rowDummy.estimate_delivery.Date <= vAvailabilityDate.Date Then
                                    Lanjut = True
                                Else
                                    Lanjut = False
                                End If
                            End If
                        Else
                            Lanjut = False
                        End If
                    ElseIf vEnumPrintAv = EnumPrintAvailability.NotAvailableOnly Then
                        If rowDummy.qty_available < rowDummy.qty_sob Then
                            Lanjut = True
                        Else
                            Lanjut = False
                        End If
                    End If
                End If

                If Lanjut = True Then
                    rowCetak = DS.tbltrans_shipping_forecast_itemdelivery.Newtbltrans_shipping_forecast_itemdeliveryRow
                    For ij As Integer = 0 To DS.tbltrans_shipping_forecast_itemdelivery.Columns.Count - 1
                        rowCetak.Item(ij) = rowDummy.Item(ij)
                    Next
                    DS.tbltrans_shipping_forecast_itemdelivery.Rows.Add(rowCetak)
                End If
            Next

            'modifikasi data
            For Each row As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow In DS.tbltrans_shipping_forecast_itemdelivery.Rows
                If row.so_number = "" Then
                    If row.IsCust_PO_NoNull = True Or row.Isso_numberNull = True Or row.Isso_dateNull = True Then
                        row.Cust_PO_No = ""
                        row.so_number = row.so_number_rec
                        row.so_date = row.so_date_rec
                    End If
                End If

                row.Cust_PO_No = row.Cust_PO_No.Trim
                If OrderNo = EnumOrderNo.Cust_PO_Number Then
                    If row.Cust_PO_No = "" Then
                        row.Cust_PO_No = row.so_number
                    End If
                ElseIf OrderNo = EnumOrderNo.Both Then
                    If row.Cust_PO_No = "" Then
                        row.Cust_PO_No = row.so_number
                    Else
                        row.Cust_PO_No = row.so_number + " (" + row.Cust_PO_No + ")"
                    End If
                Else
                    row.Cust_PO_No = row.so_number
                End If

                If Grouping = EnumPrintGrouping.Customer Then
                    row.print_detail = row.Product_Code + "(" + row.Unit_Code + ")"
                    row.print_grouping = row.customer_name
                Else
                    row.print_detail = row.customer_name
                    row.print_grouping = row.Product_Code + "(" + row.Unit_Code + ")"
                End If
            Next

            'parameter
            report.SetDataSource(DS)
            If Grouping = EnumPrintGrouping.Customer Then
                report.SetParameterValue(report.Parameter_Detail_Field.ParameterFieldName, "Code No")
                report.SetParameterValue(report.Parameter_Group_Field.ParameterFieldName, "Customer")
            Else
                report.SetParameterValue(report.Parameter_Detail_Field.ParameterFieldName, "Customer")
                report.SetParameterValue(report.Parameter_Group_Field.ParameterFieldName, "Code No")
            End If

            CRV.ReportSource = report
            CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTAHeader.Dispose()
            TBLDummy.Dispose()
        End Try
    End Function

    Public Function Cetak(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal TBLDummy As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable, _
    ByVal Grouping As EnumPrintGrouping, ByVal OrderNo As EnumOrderNo, ByVal vFilter As String, _
    ByVal vEnumPrintAv As EnumPrintAvailability, ByVal vAvailabilityDate As Date, ByVal vFilterAvDate As Boolean) As Boolean
        Dim report As New CrpTrans_MaterialPlanning_SF
        Dim DS As New DseTransaksiShippingForecast
        Dim rowDummy, rowCetak As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim DV As Data.DataRowView
        Dim Lanjut As Boolean = True

        Try
            'filtering 
            TBLDummy.DefaultView.RowFilter = vFilter
            TBLDummy.DefaultView.Sort = "[customer name] ASC, [so date rec] ASC, [so number rec] ASC, [Product Code] ASC, [promise to deliver] ASC"
            For Each DV In TBLDummy.DefaultView
                rowDummy = DV.Row
                Lanjut = True
                If Lanjut = True Then
                    If vEnumPrintAv = EnumPrintAvailability.AvailableOnly Then
                        'hanya yg available : qty available > qty sob dan tanggal available > dtp available
                        If rowDummy.qty_available >= rowDummy.qty_sob Then
                            If vFilterAvDate = False Then
                                Lanjut = True
                            Else
                                If rowDummy.estimate_delivery.Date <= vAvailabilityDate.Date Then
                                    Lanjut = True
                                Else
                                    Lanjut = False
                                End If
                            End If
                        Else
                            Lanjut = False
                        End If
                    ElseIf vEnumPrintAv = EnumPrintAvailability.NotAvailableOnly Then
                        If rowDummy.qty_available < rowDummy.qty_sob Then
                            Lanjut = True
                        Else
                            Lanjut = False
                        End If
                    End If
                End If

                If Lanjut = True Then
                    rowCetak = DS.tbltrans_shipping_forecast_itemdelivery.Newtbltrans_shipping_forecast_itemdeliveryRow
                    For ij As Integer = 0 To DS.tbltrans_shipping_forecast_itemdelivery.Columns.Count - 1
                        rowCetak.Item(ij) = rowDummy.Item(ij)
                    Next
                    DS.tbltrans_shipping_forecast_itemdelivery.Rows.Add(rowCetak)
                End If
            Next

            'modifikasi data
            For Each row As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow In DS.tbltrans_shipping_forecast_itemdelivery.Rows
                If row.so_number = "" Then
                    If row.IsCust_PO_NoNull = True Or row.Isso_numberNull = True Or row.Isso_dateNull = True Then
                        row.Cust_PO_No = ""
                        row.so_number = row.so_number_rec
                        row.so_date = row.so_date_rec
                    End If
                End If

                row.Cust_PO_No = row.Cust_PO_No.Trim
                If OrderNo = EnumOrderNo.Cust_PO_Number Then
                    If row.Cust_PO_No = "" Then
                        row.Cust_PO_No = row.so_number
                    End If
                ElseIf OrderNo = EnumOrderNo.Both Then
                    If row.Cust_PO_No = "" Then
                        row.Cust_PO_No = row.so_number
                    Else
                        row.Cust_PO_No = row.so_number + " (" + row.Cust_PO_No + ")"
                    End If
                Else
                    row.Cust_PO_No = row.so_number
                End If

                If Grouping = EnumPrintGrouping.Customer Then
                    row.print_detail = row.Product_Code + "(" + row.Unit_Code + ")"
                    row.print_grouping = row.customer_name
                Else
                    row.print_detail = row.customer_name
                    row.print_grouping = row.Product_Code + "(" + row.Unit_Code + ")"
                End If
            Next

            'parameter
            report.SetDataSource(DS)
            If Grouping = EnumPrintGrouping.Customer Then
                report.SetParameterValue(report.Parameter_Detail_Field.ParameterFieldName, "Code No")
                report.SetParameterValue(report.Parameter_Group_Field.ParameterFieldName, "Customer")
            Else
                report.SetParameterValue(report.Parameter_Detail_Field.ParameterFieldName, "Customer")
                report.SetParameterValue(report.Parameter_Group_Field.ParameterFieldName, "Code No")
            End If

            CRV.ReportSource = report
            CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Function

    Public Sub HitungEstimateDelivery1(ByVal vTBLItemSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable, _
    ByVal vIdentifier As String)
        Dim rowItemPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow
        Dim IndexItemDelivery As Integer = 0
        Dim QtyPOReady As Decimal = 0
        Dim rowStock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockRow
        Dim rowItemSO_SF As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim IndexItemPOB As Integer = 0
        Dim Dummy As Decimal = 0
        Dim DefPrec As Integer = 0
        Dim QtyReadyStock As Decimal = 0
        Dim AdaBooked As Boolean = False
        Dim vTBLItemPOB As New DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable
        Dim vTBLItemStock As New DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable
        Dim DTA_ProductQty As New DseTransaksiShippingForecastTableAdapters.tblmaster_productqtyTableAdapter
        Dim TBL_ProductQty As New DseTransaksiShippingForecast.tblmaster_productqtyDataTable
        Dim row_ProductQty As DseTransaksiShippingForecast.tblmaster_productqtyRow
        Dim DTAPendingDO As New DseTransaksiShippingForecastTableAdapters.tbltrans_do_itemTableAdapter
        Dim TBLPendingDO As New DseTransaksiShippingForecast.tbltrans_do_itemDataTable
        Dim rowPendingDO As DseTransaksiShippingForecast.tbltrans_do_itemRow
        Dim vPrimaryProduct As Double = -1
        Dim ObjSOItem As New ClsTransaksiSalesOrder
        Dim DTAPOItem As New DseTransaksiShippingForecastTableAdapters.tbltrans_po_itemTableAdapter
        Dim TBLPOItem As New DseTransaksiShippingForecast.tbltrans_po_itemDataTable
        Dim rowPOItem As DseTransaksiShippingForecast.tbltrans_po_itemRow
        Dim PrimaryPOMin, PrimaryPOMax As Double
        Dim DTAItemDeliveryPO As New DseTransaksiShippingForecastTableAdapters.tbltrans_po_itemdeliveryTableAdapter
        Dim TBLItemDeliveryPO As New DseTransaksiShippingForecast.tbltrans_po_itemdeliveryDataTable
        Dim rowItemDeliveryPO As DseTransaksiShippingForecast.tbltrans_po_itemdeliveryRow
        Dim DTAPackingList As New DseTransaksiShippingForecastTableAdapters.tblpackinglist_itemTableAdapter
        Dim TBLPackingList As New DseTransaksiShippingForecast.tblpackinglist_itemDataTable
        Dim rowPackingList As DseTransaksiShippingForecast.tblpackinglist_itemRow
        Dim DTASOItem As New DseTransaksiShippingForecastTableAdapters.tbltrans_so_itemTableAdapter
        Dim TBLSOItem As New DseTransaksiShippingForecast.tbltrans_so_itemDataTable
        Dim rowSOItem As DseTransaksiShippingForecast.tbltrans_so_itemRow
        Dim DTASOItemDelivery As New DseTransaksiShippingForecastTableAdapters.tbltrans_so_itemdeliveryTableAdapter
        Dim TBLSOItemDelivery As New DseTransaksiShippingForecast.tbltrans_so_itemdeliveryDataTable
        Dim rowSOItemDelivery As DseTransaksiShippingForecast.tbltrans_so_itemdeliveryRow
        Dim ObjPriorityLevel As New Core_Module.ClsMasterPriorityLevel
        Dim PrimaryPriorityLevelDefault As Double = ObjPriorityLevel.GetPrimaryDefault
        Dim UrutanPriodityLevelDefault As Integer = ObjPriorityLevel.GetUrutan(PrimaryPriorityLevelDefault)
        Dim DTABooking As New DseTransaksiShippingForecastTableAdapters.tbltrans_booking_invitemTableAdapter
        Dim TBLBooking As New DseTransaksiShippingForecast.tbltrans_booking_invitemDataTable
        Dim rowBooking As DseTransaksiShippingForecast.tbltrans_booking_invitemRow
        Dim IndexAcuan As Integer = 0

        Dim DTAPurchReqItem As New DseTransaksiShippingForecastTableAdapters.tbltrans_purch_req_itemTableAdapter
        Dim TBLPurchReqItem As New DseTransaksiShippingForecast.tbltrans_purch_req_itemDataTable
        Dim rowPurchReqItem As DseTransaksiShippingForecast.tbltrans_purch_req_itemRow
        Dim DTAPurchReqItemDelivery As New DseTransaksiShippingForecastTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
        Dim TBLPurchReqItemDelivery As New DseTransaksiShippingForecast.tbltrans_purchase_req_itemdeliveryDataTable
        Dim rowPurchReqItemDelivery As DseTransaksiShippingForecast.tbltrans_purchase_req_itemdeliveryRow
        Dim sdf As Double = 0

        Try
            'connection string
            DTA_ProductQty.Connection.ConnectionString = Core_Module.GetConstring
            DTAPendingDO.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTAPOItem.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTAItemDeliveryPO.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTAPackingList.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTASOItem.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTASOItemDelivery.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTABooking.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString

            vPrimaryProduct = ObjSOItem.GetPrimaryProductByIdentifier(vIdentifier)

            'isikan data dari database
            '1. stock - pending do
            DTA_ProductQty.FillBy_PrimaryProduct(TBL_ProductQty, vPrimaryProduct)
            For Each row_ProductQty In TBL_ProductQty.Rows
                rowStock = vTBLItemStock.Newtbltrans_shipping_forecast_itemstockRow
                With rowStock
                    .primary = Me.GetPrimaryMax_Stock(vTBLItemStock) + 1
                    .primary_code_no = row_ProductQty.Primary_Product
                    .primary_transaksi = -1
                    .primary_unit = row_ProductQty.Primary_Unit
                    .qty = row_ProductQty.Qty
                End With
                vTBLItemStock.Rows.Add(rowStock)
            Next
            'kurangkan data stock dengan pending do
            DTAPendingDO.FillBy_PrimaryProduct(TBLPendingDO, vPrimaryProduct)
            For Each rowPendingDO In TBLPendingDO.Rows
                vTBLItemStock.DefaultView.RowFilter = "[primary code no] = " + rowPendingDO.Primary_Product.ToString
                If vTBLItemStock.DefaultView.Count > 0 Then
                    rowStock = vTBLItemStock.DefaultView(0).Row
                    rowStock.qty -= rowPendingDO.Total_Qty
                    If rowStock.qty <= 0 Then
                        vTBLItemStock.Rows.Remove(rowStock)
                    End If
                End If
            Next

            'load data pob
            'untuk pob, untuk fungsi ini, po yang belum di otorisasi langsung di asumsi kan tidak termasuk
            'karena fungsi ini di gunakan pada form selain material planning
            'pada form material planning di asumsikan planner mengetahui mana po yang akan release, mana yang belum akan release
            'jika muncul form ini, seharusnya planner mengusahakan agar po ybs segera di otorisasi atau di remove sekalian
            DTAPOItem.FillBy_PrimaryProduct(TBLPOItem, vPrimaryProduct)
            IndexItemPOB = 0
            While IndexItemPOB < TBLPOItem.Rows.Count
                rowPOItem = TBLPOItem.Rows(IndexItemPOB)
                If rowPOItem.Authorization_Level <> rowPOItem.Authorization_Level_Needed Then
                    TBLPOItem.Rows.Remove(rowPOItem)
                Else
                    IndexItemPOB += 1
                End If
            End While

            'pindahkan ke dalam item pob
            'cari primary po max dan min
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If TBLPOItem.Rows.Count > 0 Then
                PrimaryPOMin = Double.MaxValue
                PrimaryPOMax = 0
                For Each rowPOItem In TBLPOItem.Rows
                    If PrimaryPOMax < rowPOItem.Primary_Transaksi Then
                        PrimaryPOMax = rowPOItem.Primary_Transaksi
                    End If
                    If PrimaryPOMin > rowPOItem.Primary_Transaksi Then
                        PrimaryPOMin = rowPOItem.Primary_Transaksi
                    End If
                Next
                'load item delivery
                DTAItemDeliveryPO.FillBy_PrimaryProduct(TBLItemDeliveryPO, PrimaryPOMax, PrimaryPOMin, vPrimaryProduct)
                TBLItemDeliveryPO.DefaultView.Sort = "[delivery date] DESC"
                'isikan ke dalam tabel item pob
                For Each rowPOItem In TBLPOItem.Rows
                    TBLItemDeliveryPO.DefaultView.RowFilter = "identifier = '" + rowPOItem.identifier + "'"
                    If TBLItemDeliveryPO.DefaultView.Count = 0 Then
                        'tidak ada delivery control -> set tanggal ke 1 jan 2200
                        rowItemPOB = vTBLItemPOB.Newtbltrans_shipping_forecast_item_pobRow
                        With rowItemPOB
                            .delivery_delay = 0
                            .packing_list_date = #1/1/1950#
                            .packing_list_no_rec = ""
                            .po_date_rec = rowPOItem.PO_Date
                            .po_number_rec = rowPOItem.PO_Number
                            .primary = Me.GetPrimaryMax_ItemPOB(vTBLItemPOB) + 1
                            .primary_code_no = rowPOItem.Primary_Product
                            .primary_packing_list = -1
                            .primary_po = rowPOItem.Primary_Transaksi
                            .primary_supplier = rowPOItem.Primary_Supplier
                            .primary_supplier_rec = rowPOItem.Primary_Supplier
                            .primary_transaksi = -1
                            .promise_to_deliver = #1/1/2200#
                            .qty = rowPOItem.POB
                            .remarks = ""
                            .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.POB)
                            .identifier = rowPOItem.identifier
                        End With
                        vTBLItemPOB.Rows.Add(rowItemPOB)
                    Else
                        'ambil item delivery control mulai dari yang paling besar promise to deliverynya
                        While rowPOItem.POB > 0
                            rowItemDeliveryPO = TBLItemDeliveryPO.DefaultView(0).Row
                            rowItemPOB = vTBLItemPOB.Newtbltrans_shipping_forecast_item_pobRow
                            With rowItemPOB
                                .delivery_delay = rowItemDeliveryPO.delay
                                .packing_list_date = #1/1/1950#
                                .packing_list_no_rec = ""
                                .po_date_rec = rowPOItem.PO_Date
                                .po_number_rec = rowPOItem.PO_Number
                                .primary = Me.GetPrimaryMax_ItemPOB(vTBLItemPOB) + 1
                                .primary_code_no = rowPOItem.Primary_Product
                                .primary_packing_list = -1
                                .primary_po = rowPOItem.Primary_Transaksi
                                .primary_supplier = rowPOItem.Primary_Supplier
                                .primary_supplier_rec = rowPOItem.Primary_Supplier
                                .primary_transaksi = -1
                                .promise_to_deliver = rowItemDeliveryPO.delivery_date
                                If rowPOItem.POB > rowItemDeliveryPO.delivery_qty Then
                                    .qty = rowItemDeliveryPO.delivery_qty
                                    rowPOItem.POB -= rowItemDeliveryPO.delivery_qty
                                    rowItemDeliveryPO.delivery_qty = 0
                                Else
                                    .qty = rowPOItem.POB
                                    rowPOItem.POB = 0
                                End If
                                .remarks = ""
                                .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.POB)
                                .identifier = rowPOItem.identifier
                            End With
                            vTBLItemPOB.Rows.Add(rowItemPOB)

                            'persiapan untuk looping selanjutnya
                            If rowItemDeliveryPO.delivery_qty = 0 Then
                                TBLItemDeliveryPO.Rows.Remove(rowItemDeliveryPO)
                            End If
                        End While
                    End If
                Next
            End If

            'load data git
            DTAPackingList.FillBy_PrimaryProduct(TBLPackingList, vPrimaryProduct)
            TBLPOItem.Rows.Clear()

            'isikan data delay
            If TBLPackingList.Rows.Count > 0 Then
                PrimaryPOMax = 0
                PrimaryPOMin = Double.MaxValue
                For Each rowPackingList In TBLPackingList.Rows
                    If PrimaryPOMax < rowPackingList.Primary_PO Then
                        PrimaryPOMax = rowPackingList.Primary_PO
                    End If
                    If PrimaryPOMin > rowPackingList.Primary_PO Then
                        PrimaryPOMin = rowPackingList.Primary_PO
                    End If
                Next
                DTAItemDeliveryPO.Fill(TBLItemDeliveryPO, PrimaryPOMax, PrimaryPOMin)
                For Each rowPackingList In TBLPackingList.Rows
                    TBLItemDeliveryPO.DefaultView.RowFilter = "[primary transaksi] = " + rowPackingList.Primary_PO.ToString + " AND [Primary Product] = " + rowPackingList.Primary_Product.ToString
                    If TBLItemDeliveryPO.DefaultView.Count > 0 Then
                        rowItemDeliveryPO = TBLItemDeliveryPO.DefaultView(0).Row
                        rowPackingList.delay = rowItemDeliveryPO.delay
                        rowPackingList.identifier = rowItemDeliveryPO.identifier
                    Else
                        rowPackingList.delay = 0
                        If TBLPOItem.Rows.Count = 0 Then
                            DTAPOItem.FillBy_PrimaryPOMax_Min(TBLPOItem, PrimaryPOMax, PrimaryPOMin)
                        End If
                        TBLPOItem.DefaultView.RowFilter = "[primary transaksi] = " + rowPackingList.Primary_PO.ToString + " AND [Primary Product] = " + rowPackingList.Primary_Product.ToString
                        rowPOItem = TBLPOItem.DefaultView(0).Row
                        rowPackingList.identifier = rowPOItem.identifier
                    End If
                Next
                'pindahkan ke dalam item packinglist
                For Each rowPackingList In TBLPackingList.Rows
                    rowItemPOB = vTBLItemPOB.Newtbltrans_shipping_forecast_item_pobRow
                    With rowItemPOB
                        .delivery_delay = rowPackingList.delay
                        .packing_list_date = rowPackingList.Tanggal_Transaksi
                        .packing_list_no_rec = rowPackingList.No_Transaksi
                        'untuk po date rec, diisikan dengan tanggal packing list, untuk pengurutan data ketika akan menentukan item delivery
                        .po_date_rec = #1/1/1950#
                        'untuk po number, diisikan dengan supplier packing list number, untuk record jika suatu saat diperlukan
                        .po_number_rec = rowPackingList.Supplier_Packing_List_No
                        .primary = Me.GetPrimaryMax_ItemPOB(vTBLItemPOB) + 1
                        .primary_code_no = rowPackingList.Primary_Product
                        .primary_packing_list = rowPackingList.Primary_Packing_List
                        .primary_po = rowPackingList.Primary_PO
                        .primary_supplier = rowPackingList.Primary_Supplier
                        .primary_supplier_rec = rowPackingList.Primary_Supplier
                        .primary_transaksi = -1
                        .promise_to_deliver = rowPackingList.Estimated_Arrival_Date
                        .qty = rowPackingList.Total_Qty
                        .remarks = ""
                        .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.GIT)
                        .identifier = rowPackingList.identifier
                    End With
                    vTBLItemPOB.Rows.Add(rowItemPOB)
                Next
            End If

            'sob
            DTASOItem.FillBy_PrimaryProduct(TBLSOItem, vPrimaryProduct)
            If TBLSOItem.Rows.Count > 0 Then
                PrimaryPOMax = 0
                PrimaryPOMax = Double.MinValue
                For Each rowSOItem In TBLSOItem.Rows
                    If rowSOItem.Primary_Transaksi > PrimaryPOMax Then
                        PrimaryPOMax = rowSOItem.Primary_Transaksi
                    End If
                    If PrimaryPOMin > rowSOItem.Primary_Transaksi Then
                        PrimaryPOMin = rowSOItem.Primary_Transaksi
                    End If
                Next
                DTASOItemDelivery.FillBy_Identifier2(TBLSOItemDelivery, PrimaryPOMin, PrimaryPOMax, vIdentifier)

                TBLSOItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                For Each rowSOItem In TBLSOItem.Rows
                    TBLSOItemDelivery.DefaultView.RowFilter = "identifier = '" + rowSOItem.identifier + "'"

                    If TBLSOItemDelivery.DefaultView.Count = 0 Then
                        'tambahkan ke dalam item so sesuai sob
                        rowItemSO_SF = vTBLItemSOB.Newtbltrans_shipping_forecast_itemdeliveryRow
                        With rowItemSO_SF
                            .create_do = 0
                            'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery
                            'dan 0 untuk qty available
                            .qty_available = 0
                            .qty_ready_stock = 0
                            .estimate_delivery = #1/1/2200#
                            .primary = Me.GetPrimaryMax_ItemSOB(vTBLItemSOB) + 1
                            .primary_code_no = rowSOItem.Primary_Product
                            .primary_cust = rowSOItem.Primary_Customer
                            .primary_cust_rec = rowSOItem.Primary_Customer
                            .primary_priority_level = PrimaryPriorityLevelDefault
                            .primary_so = rowSOItem.Primary_Transaksi
                            .primary_transaksi = -1
                            .primary_unit = rowSOItem.Primary_Units
                            'tidak ada delivery control -> set tanggal ke 1 jan 2200
                            .promise_to_deliver = #1/1/2200#
                            .qty_sob = rowSOItem.SOB
                            .remark = ""
                            .so_created_by = rowSOItem.Created_By
                            .so_date_rec = rowSOItem.Tanggal_Transaksi
                            .so_number_rec = rowSOItem.No_Transaksi
                            .urutan_priority_level = UrutanPriodityLevelDefault
                            .customer_name = rowSOItem.Customer_Name
                            .Product_Code = rowSOItem.Product_Code
                            .Description = rowSOItem.Description
                            .so_number = rowSOItem.No_Transaksi
                            .so_date = rowSOItem.Tanggal_Transaksi
                            .Cust_PO_No = rowSOItem.Cust_PO_No
                            ._Decimal = rowSOItem._Decimal
                            .Unit_Code = rowSOItem.Unit_Code
                            .Primary_Category = rowSOItem.Primary_Category
                            .Primary_Sub_Brand = rowSOItem.Primary_Sub_Brand
                            .Primary_Brand = rowSOItem.Primary_Brand
                            .primary_product_group = rowSOItem.Primary_Product_Group
                            .identifier = rowSOItem.identifier
                            .booked = 0
                            .no_booking = ""
                            .qty_booked = 0
                            .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB)
                            .tipe_product = rowSOItem.Type_Product
                        End With
                        vTBLItemSOB.Rows.Add(rowItemSO_SF)
                    Else
                        'ambil item delivery control mulai dari yang paling besar promise to deliverynya
                        While rowSOItem.SOB > 0
                            rowSOItemDelivery = TBLSOItemDelivery.DefaultView(0).Row
                            rowItemSO_SF = vTBLItemSOB.Newtbltrans_shipping_forecast_itemdeliveryRow
                            With rowItemSO_SF
                                .create_do = 0
                                'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery dan 0 untuk qty availabel
                                .estimate_delivery = #1/1/2200#
                                .qty_available = 0
                                .qty_ready_stock = 0
                                .primary = Me.GetPrimaryMax_ItemSOB(vTBLItemSOB) + 1
                                .primary_code_no = rowSOItem.Primary_Product
                                .primary_cust = rowSOItem.Primary_Customer
                                .primary_cust_rec = rowSOItem.Primary_Customer
                                .primary_priority_level = rowSOItemDelivery.delivery_priority
                                .urutan_priority_level = rowSOItemDelivery.urutan
                                .primary_so = rowSOItem.Primary_Transaksi
                                .primary_transaksi = -1
                                .primary_unit = rowSOItem.Primary_Units
                                'tidak ada delivery control -> set tanggal ke 1 jan 2200
                                .promise_to_deliver = rowSOItemDelivery.delivery_date
                                If rowSOItem.SOB > rowSOItemDelivery.delivery_qty Then
                                    .qty_sob = rowSOItemDelivery.delivery_qty
                                    rowSOItem.SOB -= rowSOItemDelivery.delivery_qty
                                    rowSOItemDelivery.delivery_qty = 0
                                Else
                                    .qty_sob = rowSOItem.SOB
                                    rowSOItem.SOB = 0
                                End If
                                .remark = ""
                                .so_created_by = rowSOItem.Created_By
                                .so_date_rec = rowSOItem.Tanggal_Transaksi
                                .so_number_rec = rowSOItem.No_Transaksi
                                .customer_name = rowSOItem.Customer_Name
                                .Product_Code = rowSOItem.Product_Code
                                .Description = rowSOItem.Description
                                .so_number = rowSOItem.No_Transaksi
                                .so_date = rowSOItem.Tanggal_Transaksi
                                .Cust_PO_No = rowSOItem.Cust_PO_No
                                ._Decimal = rowSOItem._Decimal
                                .Unit_Code = rowSOItem.Unit_Code
                                .Primary_Category = rowSOItem.Primary_Category
                                .Primary_Sub_Brand = rowSOItem.Primary_Sub_Brand
                                .Primary_Brand = rowSOItem.Primary_Brand
                                .primary_product_group = rowSOItem.Primary_Product_Group
                                .identifier = rowSOItem.identifier
                                .booked = 0
                                .no_booking = ""
                                .qty_booked = 0
                                .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB)
                                .tipe_product = rowSOItem.Type_Product
                                If rowSOItemDelivery.delivery_qty = 0 Then
                                    TBLSOItemDelivery.Rows.Remove(rowSOItemDelivery)
                                End If
                            End With
                            vTBLItemSOB.Rows.Add(rowItemSO_SF)
                        End While
                    End If
                Next
            End If

            'tandain item sob yang ada inventory bookingnya
            DTABooking.FillBy_PrimaryCodeNo(TBLBooking, vPrimaryProduct)
            For Each rowBooking In TBLBooking.Rows
                vTBLItemSOB.DefaultView.RowFilter = "[primary code no] = " + rowBooking.Primary_Product.ToString + " AND [primary so] = " + rowBooking.Primary_SO.ToString
                vTBLItemSOB.DefaultView.Sort = "[promise to deliver] ASC, [urutan priority level] ASC, [so date rec] ASC, [primary so] ASC"
                If vTBLItemSOB.DefaultView.Count > 0 Then
                    IndexAcuan = 0
                    While IndexAcuan < vTBLItemSOB.DefaultView.Count And rowBooking.Qty > 0
                        rowItemSO_SF = vTBLItemSOB.DefaultView(IndexAcuan).Row

                        If rowItemSO_SF.qty_sob - rowItemSO_SF.qty_booked >= rowBooking.Qty Then
                            rowItemSO_SF.qty_booked += rowBooking.Qty
                            rowBooking.Qty = 0
                        Else
                            rowBooking.Qty -= (rowItemSO_SF.qty_sob - rowItemSO_SF.qty_booked)
                            rowItemSO_SF.qty_booked = rowItemSO_SF.qty_sob
                            IndexAcuan += 1
                        End If

                        rowItemSO_SF.booked = 1
                        If rowItemSO_SF.no_booking = "" Then
                            rowItemSO_SF.no_booking = rowBooking.No_Transaksi
                        Else
                            If rowItemSO_SF.no_booking.Contains(rowBooking.No_Transaksi) = False Then
                                rowItemSO_SF.no_booking = rowItemSO_SF.no_booking + ", " + rowBooking.No_Transaksi
                            End If
                        End If
                    End While
                End If
            Next

            'tambahkan item sob dengan purchase request
            DTAPurchReqItem.Fill(TBLPurchReqItem)
            If TBLPurchReqItem.Rows.Count > 0 Then
                PrimaryPOMax = 0
                PrimaryPOMax = Double.MinValue
                For Each rowPurchReqItem In TBLPurchReqItem.Rows
                    If rowPurchReqItem.primary_transaksi > PrimaryPOMax Then
                        PrimaryPOMax = rowPurchReqItem.primary_transaksi
                    End If
                    If PrimaryPOMin > rowPurchReqItem.primary_transaksi Then
                        PrimaryPOMin = rowPurchReqItem.primary_transaksi
                    End If
                Next
                DTAPurchReqItemDelivery.Fill(TBLPurchReqItemDelivery, PrimaryPOMax, PrimaryPOMin)

                TBLPurchReqItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                For Each rowPurchReqItem In TBLPurchReqItem.Rows
                    TBLPurchReqItemDelivery.DefaultView.RowFilter = "identifier = '" + rowPurchReqItem.identifier + "'"

                    If TBLPurchReqItemDelivery.DefaultView.Count = 0 Then
                        'tambahkan ke dalam item sob sesuai Purc Req OB
                        rowItemSO_SF = vTBLItemSOB.Newtbltrans_shipping_forecast_itemdeliveryRow
                        With rowItemSO_SF
                            .create_do = 0
                            'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery
                            'dan 0 untuk qty available
                            .qty_available = 0
                            .qty_ready_stock = 0
                            .estimate_delivery = #1/1/2200#
                            .primary = Me.GetPrimaryMax_ItemSOB(vTBLItemSOB) + 1
                            .primary_code_no = rowPurchReqItem.primary_code_no
                            .primary_cust = rowPurchReqItem.primary_ref_party
                            .primary_cust_rec = rowPurchReqItem.primary_ref_party
                            .primary_priority_level = PrimaryPriorityLevelDefault
                            .primary_so = rowPurchReqItem.primary_transaksi
                            .primary_transaksi = -1
                            .primary_unit = rowPurchReqItem.primary_unit
                            'tidak ada delivery control -> set tanggal ke 1 jan 2200
                            .promise_to_deliver = #1/1/2200#
                            .qty_sob = rowPurchReqItem.OB
                            .remark = rowPurchReqItem.remarks_item
                            .so_created_by = rowPurchReqItem.created_by
                            .so_date_rec = rowPurchReqItem.tanggal
                            .so_number_rec = rowPurchReqItem.no_transaksi
                            .urutan_priority_level = UrutanPriodityLevelDefault
                            .customer_name = rowPurchReqItem.Ref_Party_Desc
                            .Product_Code = rowPurchReqItem.Product_Code
                            .Description = rowPurchReqItem.Description
                            .so_number = rowPurchReqItem.no_transaksi
                            .so_date = rowPurchReqItem.tanggal
                            .Cust_PO_No = ""
                            ._Decimal = rowPurchReqItem._Decimal
                            .Unit_Code = rowPurchReqItem.Unit_Code
                            .Primary_Category = rowPurchReqItem.Primary_Category
                            .Primary_Sub_Brand = rowPurchReqItem.Primary_Sub_Brand
                            .Primary_Brand = rowPurchReqItem.Primary_Brand
                            .primary_product_group = rowPurchReqItem.Primary_Product_Group
                            .identifier = rowPurchReqItem.identifier
                            .booked = 0
                            .no_booking = ""
                            .qty_booked = 0
                            .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req)
                            .tipe_product = rowPurchReqItem.Type
                        End With
                        vTBLItemSOB.Rows.Add(rowItemSO_SF)
                    Else
                        'ambil item delivery control mulai dari yang paling besar promise to deliverynya
                        While rowPurchReqItem.OB > 0
                            rowPurchReqItemDelivery = TBLPurchReqItemDelivery.DefaultView(0).Row
                            rowItemSO_SF = vTBLItemSOB.Newtbltrans_shipping_forecast_itemdeliveryRow
                            With rowItemSO_SF
                                .create_do = 0
                                'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery dan 0 untuk qty availabel
                                .estimate_delivery = #1/1/2200#
                                .qty_available = 0
                                .qty_ready_stock = 0
                                .primary = Me.GetPrimaryMax_ItemSOB(vTBLItemSOB) + 1
                                .primary_code_no = rowPurchReqItem.primary_code_no
                                .primary_cust = rowPurchReqItem.primary_ref_party
                                .primary_cust_rec = rowPurchReqItem.primary_ref_party
                                .primary_priority_level = rowPurchReqItemDelivery.delivery_priority
                                .urutan_priority_level = rowPurchReqItemDelivery.urutan
                                .primary_so = rowPurchReqItem.primary_transaksi
                                .primary_transaksi = -1
                                .primary_unit = rowPurchReqItem.primary_unit
                                'tidak ada delivery control -> set tanggal ke 1 jan 2200
                                .promise_to_deliver = rowPurchReqItemDelivery.delivery_date
                                If rowPurchReqItem.OB > rowPurchReqItemDelivery.delivery_qty Then
                                    .qty_sob = rowPurchReqItemDelivery.delivery_qty
                                    rowPurchReqItem.OB -= rowPurchReqItemDelivery.delivery_qty
                                    rowPurchReqItemDelivery.delivery_qty = 0
                                Else
                                    .qty_sob = rowPurchReqItem.OB
                                    rowPurchReqItem.OB = 0
                                End If
                                .remark = rowPurchReqItem.remarks_item
                                .so_created_by = rowPurchReqItem.created_by
                                .so_date_rec = rowPurchReqItem.tanggal
                                .so_number_rec = rowPurchReqItem.no_transaksi
                                .customer_name = rowPurchReqItem.Ref_Party_Desc
                                .Product_Code = rowPurchReqItem.Product_Code
                                .Description = rowPurchReqItem.Description
                                .so_number = rowPurchReqItem.no_transaksi
                                .so_date = rowPurchReqItem.tanggal
                                .Cust_PO_No = ""
                                ._Decimal = rowPurchReqItem._Decimal
                                .Unit_Code = rowPurchReqItem.Unit_Code
                                .Primary_Category = rowPurchReqItem.Primary_Category
                                .Primary_Sub_Brand = rowPurchReqItem.Primary_Sub_Brand
                                .Primary_Brand = rowPurchReqItem.Primary_Brand
                                .primary_product_group = rowPurchReqItem.Primary_Product_Group
                                .identifier = rowPurchReqItem.identifier
                                .booked = 0
                                .no_booking = ""
                                .qty_booked = 0
                                .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req)
                                .tipe_product = rowPurchReqItem.Type
                                If rowPurchReqItemDelivery.delivery_qty = 0 Then
                                    TBLPurchReqItemDelivery.Rows.Remove(rowPurchReqItemDelivery)
                                End If
                            End With
                            vTBLItemSOB.Rows.Add(rowItemSO_SF)
                        End While
                    End If
                Next
            End If

            'menghitung perkiraan kapan pengiriman selanjutnya
            For Each rowItemPOB In vTBLItemPOB.Rows
                'set parameter item pob yang akan digunakan pada perhitungan delivery date
                rowItemPOB.available_date = rowItemPOB.promise_to_deliver + TimeSpan.FromDays(rowItemPOB.delivery_delay)
                rowItemPOB.qty_used = 0
            Next
            'set parameter item sob yang akan digunakan pada perhitungan delivery date
            For Each rowItemSO_SF In vTBLItemSOB.Rows
                If rowItemSO_SF.tipe_product = Core_Module.clsMasterProduct.EnumTipeProduct.Services Then
                    'karena ready stock, maka qty availabel = qty sob
                    rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
                    rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                    'karena ready stock, maka tandai estimate delivery menjadi ASAP -> 1 jan 1950
                    rowItemSO_SF.estimate_delivery = #1/1/1950#
                Else
                    rowItemSO_SF.qty_available = 0
                    rowItemSO_SF.estimate_delivery = #1/1/2200#
                    rowItemSO_SF.qty_ready_stock = 0
                End If
            Next

            'set parameter item stock untuk perhitungan ulang
            For Each rowStock In vTBLItemStock.Rows
                rowStock.qty_used = 0
            Next

            'filter + sort
            vTBLItemPOB.DefaultView.Sort = "[available date] ASC, [promise to deliver] ASC, [po date rec] ASC, [primary po] ASC"
            vTBLItemSOB.DefaultView.Sort = "[booked] DESC, [promise to deliver] ASC, [urutan priority level] ASC, [so date rec] ASC, [primary so] ASC"

            'kembalikan parameter ke nilai awal
            IndexItemDelivery = 0
            QtyPOReady = 0
            QtyReadyStock = 0

            'pasangkan dengan yang ready stock dahulu
            If vTBLItemStock.DefaultView.Count > 0 Then
                rowStock = vTBLItemStock.DefaultView(0).Row

                While rowStock.qty - rowStock.qty_used > 0 And vTBLItemSOB.DefaultView.Count > IndexItemDelivery
                    rowItemSO_SF = vTBLItemSOB.DefaultView(IndexItemDelivery).Row
                    If rowStock.qty - rowStock.qty_used >= rowItemSO_SF.qty_sob Then
                        rowStock.qty_used += rowItemSO_SF.qty_sob
                        'karena ready stock, maka qty availabel = qty sob dan qty ready stock = qty sob
                        rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
                        rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                        'karena ready stock, maka tandai estimate delivery menjadi ASAP -> 1 jan 1950
                        rowItemSO_SF.estimate_delivery = #1/1/1950#
                        'naikan index item delivery, cek item delivery selanjutnya
                        IndexItemDelivery += 1
                    Else
                        QtyPOReady = rowStock.qty - rowStock.qty_used
                        rowStock.qty_used = rowStock.qty
                        QtyReadyStock = QtyPOReady
                        Exit While
                    End If
                End While
            End If

            'pasangkan dengan item po
            If vTBLItemPOB.DefaultView.Count > 0 Then
                IndexItemPOB = 0
                While IndexItemPOB < vTBLItemPOB.DefaultView.Count And IndexItemDelivery < vTBLItemSOB.DefaultView.Count
                    rowItemPOB = vTBLItemPOB.DefaultView(IndexItemPOB).Row
                    rowItemSO_SF = vTBLItemSOB.DefaultView(IndexItemDelivery).Row
                    If QtyPOReady + rowItemPOB.qty - rowItemPOB.qty_used >= rowItemSO_SF.qty_sob Then
                        'jika item yang tersedia lebih besar dari item delivery, maka 
                        'set delivery date + qty available delivery + pindah ke baris item delivery selanjutnya
                        rowItemSO_SF.estimate_delivery = rowItemPOB.available_date
                        rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
                        'majukan index item delivery
                        IndexItemDelivery += 1

                        'kurangkan qty po ready
                        Dummy = rowItemSO_SF.qty_sob
                        If QtyPOReady >= Dummy Then
                            QtyPOReady -= Dummy
                            Dummy = 0
                        Else
                            Dummy -= QtyPOReady
                            QtyPOReady = 0
                        End If

                        'tambahkan qty used pada item po, untuk menanda jumlah yang sudah digunakan
                        rowItemPOB.qty_used += Dummy

                        'jika semua qty sudah habis dipakai, maka pindah ke baris item pob selanjutnya
                        If rowItemPOB.qty <= rowItemPOB.qty_used Then
                            IndexItemPOB += 1
                        End If

                        'perhitungan qty ready stock 
                        If QtyReadyStock >= rowItemSO_SF.qty_sob Then
                            rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                            QtyReadyStock -= rowItemSO_SF.qty_sob
                        Else
                            rowItemSO_SF.qty_ready_stock = QtyReadyStock
                            QtyReadyStock = 0
                        End If
                    Else
                        'jika item delivery lebih besar dari qty po ready + qty po,ada 2 kemungkinan:
                        If IndexItemPOB + 1 < vTBLItemPOB.DefaultView.Count Then
                            '1. tambahkan qty po dari baris pob selanjut, karena dari pob, maka qty ready stock tidak di tambahkan
                            QtyPOReady += rowItemPOB.qty - rowItemPOB.qty_used
                            rowItemPOB.qty_used = rowItemPOB.qty
                            IndexItemPOB += 1
                        Else
                            '2. tidak ada baris pob selanjutnya, maka qty availble hanya sejumlah yang ada,
                            'estimate delivery tidak dikehui kapan, tapi available qty ada
                            rowItemSO_SF.qty_available = rowItemPOB.qty - rowItemPOB.qty_used + QtyPOReady
                            rowItemSO_SF.estimate_delivery = #1/1/2200#
                            'po ready di nol kan karena sudah habis
                            QtyPOReady = 0
                            'perhitungan ready stock
                            If QtyReadyStock >= rowItemSO_SF.qty_sob Then
                                rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                                QtyReadyStock -= rowItemSO_SF.qty_sob
                            Else
                                rowItemSO_SF.qty_ready_stock = QtyReadyStock
                                QtyReadyStock = 0
                            End If
                            'tambahkan baris dgv
                            'tambahan index item delivery, untuk sisanya
                            IndexItemDelivery += 1
                            'index item pob tetap dinaikan, untuk menghentikan looping perhitungan
                            IndexItemPOB += 1
                        End If
                    End If
                End While
            End If

            'jika masih ada sisa item pob yang belum di handle, maka tambahkan ke dgv
            IndexItemPOB += 1
            While IndexItemPOB < vTBLItemPOB.DefaultView.Count
                rowItemPOB = vTBLItemPOB.DefaultView(IndexItemPOB).Row
                IndexItemPOB += 1
            End While
            'jika masih ada sisa item sob yang belum di handle, maka tambahkan ke dgv
            While IndexItemDelivery < vTBLItemSOB.DefaultView.Count
                rowItemSO_SF = vTBLItemSOB.DefaultView(IndexItemDelivery).Row
                If QtyPOReady > 0 Then
                    rowItemSO_SF.qty_available += QtyPOReady
                    QtyPOReady = 0
                End If
                If QtyReadyStock > 0 Then
                    rowItemSO_SF.qty_ready_stock += QtyReadyStock
                    QtyReadyStock = 0
                End If
                IndexItemDelivery += 1
            End While

            'Buang untuk item sob yang tidak sesuai dengan identifier yang diberikan
            IndexAcuan = 0
            While IndexAcuan < vTBLItemSOB.Rows.Count
                rowItemSO_SF = vTBLItemSOB.Rows(IndexAcuan)
                If rowItemSO_SF.identifier = vIdentifier And rowItemSO_SF.tipe = ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB Then
                    IndexAcuan += 1
                Else
                    vTBLItemSOB.Rows.Remove(rowItemSO_SF)
                End If
            End While

            'untuk yang service set ready
            For Each rowItemSO_SF In vTBLItemSOB.Rows
                If rowItemSO_SF.tipe_product = Core_Module.clsMasterProduct.EnumTipeProduct.Services Then
                    'karena ready stock, maka qty availabel = qty sob
                    rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
                    rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                    'karena ready stock, maka tandai estimate delivery menjadi ASAP -> 1 jan 1950
                    rowItemSO_SF.estimate_delivery = #1/1/1950#
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vTBLItemPOB.Dispose()
            DTA_ProductQty.Dispose()
            TBL_ProductQty.Dispose()
            DTAPendingDO.Dispose()
            TBLPendingDO.Dispose()
            ObjSOItem = Nothing
            DTAPOItem.Dispose()
            TBLPOItem.Dispose()
            DTAItemDeliveryPO.Dispose()
            TBLItemDeliveryPO.Dispose()
            DTAPackingList.Dispose()
            TBLPackingList.Dispose()
            DTASOItem.Dispose()
            TBLSOItem.Dispose()
            DTASOItemDelivery.Dispose()
            TBLSOItemDelivery.Dispose()
            ObjPriorityLevel = Nothing
            DTABooking.Dispose()
            TBLBooking.Dispose()

            DTAPurchReqItem.Dispose()
            TBLPurchReqItem.Dispose()
            DTAPurchReqItemDelivery.Dispose()
            TBLPurchReqItemDelivery.Dispose()
        End Try
    End Sub

    ' Public Sub HitungEstimateDelivery1(ByVal vTBLItemSOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable, _
    'ByVal vPrimaryProduct As Double)
    '     Dim rowItemPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow
    '     Dim IndexItemDelivery As Integer = 0
    '     Dim QtyPOReady As Decimal = 0
    '     Dim rowStock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockRow
    '     Dim rowItemSO_SF As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
    '     Dim IndexItemPOB As Integer = 0
    '     Dim Dummy As Decimal = 0
    '     Dim DefPrec As Integer = 0
    '     Dim QtyReadyStock As Decimal = 0
    '     Dim AdaBooked As Boolean = False
    '     Dim vTBLItemPOB As New DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable
    '     Dim vTBLItemStock As New DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable
    '     Dim DTA_ProductQty As New DseTransaksiShippingForecastTableAdapters.tblmaster_productqtyTableAdapter
    '     Dim TBL_ProductQty As New DseTransaksiShippingForecast.tblmaster_productqtyDataTable
    '     Dim row_ProductQty As DseTransaksiShippingForecast.tblmaster_productqtyRow
    '     Dim DTAPendingDO As New DseTransaksiShippingForecastTableAdapters.tbltrans_do_itemTableAdapter
    '     Dim TBLPendingDO As New DseTransaksiShippingForecast.tbltrans_do_itemDataTable
    '     Dim rowPendingDO As DseTransaksiShippingForecast.tbltrans_do_itemRow
    '     Dim ObjSOItem As New ClsTransaksiSalesOrder
    '     Dim DTAPOItem As New DseTransaksiShippingForecastTableAdapters.tbltrans_po_itemTableAdapter
    '     Dim TBLPOItem As New DseTransaksiShippingForecast.tbltrans_po_itemDataTable
    '     Dim rowPOItem As DseTransaksiShippingForecast.tbltrans_po_itemRow
    '     Dim PrimaryPOMin, PrimaryPOMax As Double
    '     Dim DTAItemDeliveryPO As New DseTransaksiShippingForecastTableAdapters.tbltrans_po_itemdeliveryTableAdapter
    '     Dim TBLItemDeliveryPO As New DseTransaksiShippingForecast.tbltrans_po_itemdeliveryDataTable
    '     Dim rowItemDeliveryPO As DseTransaksiShippingForecast.tbltrans_po_itemdeliveryRow
    '     Dim DTAPackingList As New DseTransaksiShippingForecastTableAdapters.tblpackinglist_itemTableAdapter
    '     Dim TBLPackingList As New DseTransaksiShippingForecast.tblpackinglist_itemDataTable
    '     Dim rowPackingList As DseTransaksiShippingForecast.tblpackinglist_itemRow
    '     Dim DTASOItem As New DseTransaksiShippingForecastTableAdapters.tbltrans_so_itemTableAdapter
    '     Dim TBLSOItem As New DseTransaksiShippingForecast.tbltrans_so_itemDataTable
    '     Dim rowSOItem As DseTransaksiShippingForecast.tbltrans_so_itemRow
    '     Dim DTASOItemDelivery As New DseTransaksiShippingForecastTableAdapters.tbltrans_so_itemdeliveryTableAdapter
    '     Dim TBLSOItemDelivery As New DseTransaksiShippingForecast.tbltrans_so_itemdeliveryDataTable
    '     Dim rowSOItemDelivery As DseTransaksiShippingForecast.tbltrans_so_itemdeliveryRow
    '     Dim ObjPriorityLevel As New Core_Module.ClsMasterPriorityLevel
    '     Dim PrimaryPriorityLevelDefault As Double = ObjPriorityLevel.GetPrimaryDefault
    '     Dim UrutanPriodityLevelDefault As Integer = ObjPriorityLevel.GetUrutan(PrimaryPriorityLevelDefault)
    '     Dim DTABooking As New DseTransaksiShippingForecastTableAdapters.tbltrans_booking_invitemTableAdapter
    '     Dim TBLBooking As New DseTransaksiShippingForecast.tbltrans_booking_invitemDataTable
    '     Dim rowBooking As DseTransaksiShippingForecast.tbltrans_booking_invitemRow
    '     Dim IndexAcuan As Integer = 0

    '     Try
    '         'connection string
    '         DTA_ProductQty.Connection.ConnectionString = Core_Module.GetConstring
    '         DTAPendingDO.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
    '         DTAPOItem.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
    '         DTAItemDeliveryPO.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
    '         DTAPackingList.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
    '         DTASOItem.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
    '         DTASOItemDelivery.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
    '         DTABooking.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString

    '         'isikan data dari database
    '         '1. stock - pending do
    '         DTA_ProductQty.FillBy_PrimaryProduct(TBL_ProductQty, vPrimaryProduct)
    '         For Each row_ProductQty In TBL_ProductQty.Rows
    '             rowStock = vTBLItemStock.Newtbltrans_shipping_forecast_itemstockRow
    '             With rowStock
    '                 .primary = Me.GetPrimaryMax_Stock(vTBLItemStock) + 1
    '                 .primary_code_no = row_ProductQty.Primary_Product
    '                 .primary_transaksi = -1
    '                 .primary_unit = row_ProductQty.Primary_Unit
    '                 .qty = row_ProductQty.Qty
    '             End With
    '             vTBLItemStock.Rows.Add(rowStock)
    '         Next
    '         'kurangkan data stock dengan pending do
    '         DTAPendingDO.FillBy_PrimaryProduct(TBLPendingDO, vPrimaryProduct)
    '         For Each rowPendingDO In TBLPendingDO.Rows
    '             vTBLItemStock.DefaultView.RowFilter = "[primary code no] = " + rowPendingDO.Primary_Product.ToString
    '             If vTBLItemStock.DefaultView.Count > 0 Then
    '                 rowStock = vTBLItemStock.DefaultView(0).Row
    '                 rowStock.qty -= rowPendingDO.Total_Qty
    '                 If rowStock.qty <= 0 Then
    '                     vTBLItemStock.Rows.Remove(rowStock)
    '                 End If
    '             End If
    '         Next

    '         'load data pob
    '         'untuk pob, untuk fungsi ini, po yang belum di otorisasi langsung di asumsi kan tidak termasuk
    '         'karena fungsi ini di gunakan pada form selain material planning
    '         'pada form material planning di asumsikan planner mengetahui mana po yang akan release, mana yang belum akan release
    '         'jika muncul form ini, seharusnya planner mengusahakan agar po ybs segera di otorisasi atau di remove sekalian
    '         DTAPOItem.FillBy_PrimaryProduct(TBLPOItem, vPrimaryProduct)
    '         IndexItemPOB = 0
    '         While IndexItemPOB < TBLPOItem.Rows.Count
    '             rowPOItem = TBLPOItem.Rows(IndexItemPOB)
    '             If rowPOItem.Authorization_Level <> rowPOItem.Authorization_Level_Needed Then
    '                 TBLPOItem.Rows.Remove(rowPOItem)
    '             Else
    '                 IndexItemPOB += 1
    '             End If
    '         End While

    '         'pindahkan ke dalam item pob
    '         'cari primary po max dan min
    '         Windows.Forms.Cursor.Current = Cursors.WaitCursor
    '         If TBLPOItem.Rows.Count > 0 Then
    '             PrimaryPOMin = Double.MaxValue
    '             PrimaryPOMax = 0
    '             For Each rowPOItem In TBLPOItem.Rows
    '                 If PrimaryPOMax < rowPOItem.Primary_Transaksi Then
    '                     PrimaryPOMax = rowPOItem.Primary_Transaksi
    '                 End If
    '                 If PrimaryPOMin > rowPOItem.Primary_Transaksi Then
    '                     PrimaryPOMin = rowPOItem.Primary_Transaksi
    '                 End If
    '             Next
    '             'load item delivery
    '             DTAItemDeliveryPO.FillBy_PrimaryProduct(TBLItemDeliveryPO, PrimaryPOMax, PrimaryPOMin, vPrimaryProduct)
    '             TBLItemDeliveryPO.DefaultView.Sort = "[delivery date] DESC"
    '             'isikan ke dalam tabel item pob
    '             For Each rowPOItem In TBLPOItem.Rows
    '                 TBLItemDeliveryPO.DefaultView.RowFilter = "identifier = '" + rowPOItem.identifier + "'"
    '                 If TBLItemDeliveryPO.DefaultView.Count = 0 Then
    '                     'tidak ada delivery control -> set tanggal ke 1 jan 2200
    '                     rowItemPOB = vTBLItemPOB.Newtbltrans_shipping_forecast_item_pobRow
    '                     With rowItemPOB
    '                         .delivery_delay = 0
    '                         .packing_list_date = #1/1/1950#
    '                         .packing_list_no_rec = ""
    '                         .po_date_rec = rowPOItem.PO_Date
    '                         .po_number_rec = rowPOItem.PO_Number
    '                         .primary = Me.GetPrimaryMax_ItemPOB(vTBLItemPOB) + 1
    '                         .primary_code_no = rowPOItem.Primary_Product
    '                         .primary_packing_list = -1
    '                         .primary_po = rowPOItem.Primary_Transaksi
    '                         .primary_supplier = rowPOItem.Primary_Supplier
    '                         .primary_supplier_rec = rowPOItem.Primary_Supplier
    '                         .primary_transaksi = -1
    '                         .promise_to_deliver = #1/1/2200#
    '                         .qty = rowPOItem.POB
    '                         .remarks = ""
    '                         .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.POB)
    '                         .identifier = rowPOItem.identifier
    '                     End With
    '                     vTBLItemPOB.Rows.Add(rowItemPOB)
    '                 Else
    '                     'ambil item delivery control mulai dari yang paling besar promise to deliverynya
    '                     While rowPOItem.POB > 0
    '                         rowItemDeliveryPO = TBLItemDeliveryPO.DefaultView(0).Row
    '                         rowItemPOB = vTBLItemPOB.Newtbltrans_shipping_forecast_item_pobRow
    '                         With rowItemPOB
    '                             .delivery_delay = rowItemDeliveryPO.delay
    '                             .packing_list_date = #1/1/1950#
    '                             .packing_list_no_rec = ""
    '                             .po_date_rec = rowPOItem.PO_Date
    '                             .po_number_rec = rowPOItem.PO_Number
    '                             .primary = Me.GetPrimaryMax_ItemPOB(vTBLItemPOB) + 1
    '                             .primary_code_no = rowPOItem.Primary_Product
    '                             .primary_packing_list = -1
    '                             .primary_po = rowPOItem.Primary_Transaksi
    '                             .primary_supplier = rowPOItem.Primary_Supplier
    '                             .primary_supplier_rec = rowPOItem.Primary_Supplier
    '                             .primary_transaksi = -1
    '                             .promise_to_deliver = rowItemDeliveryPO.delivery_date
    '                             If rowPOItem.POB > rowItemDeliveryPO.delivery_qty Then
    '                                 .qty = rowItemDeliveryPO.delivery_qty
    '                                 rowPOItem.POB -= rowItemDeliveryPO.delivery_qty
    '                                 rowItemDeliveryPO.delivery_qty = 0
    '                             Else
    '                                 .qty = rowPOItem.POB
    '                                 rowPOItem.POB = 0
    '                             End If
    '                             .remarks = ""
    '                             .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.POB)
    '                             .identifier = rowPOItem.identifier
    '                         End With
    '                         vTBLItemPOB.Rows.Add(rowItemPOB)

    '                         'persiapan untuk looping selanjutnya
    '                         If rowItemDeliveryPO.delivery_qty = 0 Then
    '                             TBLItemDeliveryPO.Rows.Remove(rowItemDeliveryPO)
    '                         End If
    '                     End While
    '                 End If
    '             Next
    '         End If

    '         'load data git
    '         DTAPackingList.FillBy_PrimaryProduct(TBLPackingList, vPrimaryProduct)
    '         'isikan data delay
    '         If TBLPackingList.Rows.Count > 0 Then
    '             PrimaryPOMax = 0
    '             PrimaryPOMin = Double.MaxValue
    '             For Each rowPackingList In TBLPackingList.Rows
    '                 If PrimaryPOMax < rowPackingList.Primary_PO Then
    '                     PrimaryPOMax = rowPackingList.Primary_PO
    '                 End If
    '                 If PrimaryPOMin > rowPackingList.Primary_PO Then
    '                     PrimaryPOMin = rowPackingList.Primary_PO
    '                 End If
    '             Next
    '             DTAItemDeliveryPO.Fill(TBLItemDeliveryPO, PrimaryPOMax, PrimaryPOMin)
    '             TBLPOItem.Rows.Clear()
    '             For Each rowPackingList In TBLPackingList.Rows
    '                 TBLItemDeliveryPO.DefaultView.RowFilter = "[primary transaksi] = " + rowPackingList.Primary_PO.ToString + " AND [Primary Product] = " + rowPackingList.Primary_Product.ToString
    '                 If TBLItemDeliveryPO.DefaultView.Count > 0 Then
    '                     rowItemDeliveryPO = TBLItemDeliveryPO.DefaultView(0).Row
    '                     rowPackingList.delay = rowItemDeliveryPO.delay
    '                     rowPackingList.identifier = rowItemDeliveryPO.identifier
    '                 Else
    '                     rowPackingList.delay = 0
    '                     If TBLPOItem.Rows.Count = 0 Then
    '                         DTAPOItem.FillBy_PrimaryPOMax_Min(TBLPOItem, PrimaryPOMax, PrimaryPOMin)
    '                     End If
    '                     TBLPOItem.DefaultView.RowFilter = "[primary transaksi] = " + rowPackingList.Primary_PO.ToString + " AND [Primary Product] = " + rowPackingList.Primary_Product.ToString
    '                     rowPOItem = TBLPOItem.DefaultView(0).Row
    '                     rowPackingList.identifier = rowPOItem.identifier
    '                 End If
    '             Next
    '             'pindahkan ke dalam item packinglist
    '             For Each rowPackingList In TBLPackingList.Rows
    '                 rowItemPOB = vTBLItemPOB.Newtbltrans_shipping_forecast_item_pobRow
    '                 With rowItemPOB
    '                     .delivery_delay = rowPackingList.delay
    '                     .packing_list_date = rowPackingList.Tanggal_Transaksi
    '                     .packing_list_no_rec = rowPackingList.No_Transaksi
    '                     'untuk po date rec, diisikan dengan tanggal packing list, untuk pengurutan data ketika akan menentukan item delivery
    '                     .po_date_rec = #1/1/1950#
    '                     'untuk po number, diisikan dengan supplier packing list number, untuk record jika suatu saat diperlukan
    '                     .po_number_rec = rowPackingList.Supplier_Packing_List_No
    '                     .primary = Me.GetPrimaryMax_ItemPOB(vTBLItemPOB) + 1
    '                     .primary_code_no = rowPackingList.Primary_Product
    '                     .primary_packing_list = rowPackingList.Primary_Packing_List
    '                     .primary_po = rowPackingList.Primary_PO
    '                     .primary_supplier = rowPackingList.Primary_Supplier
    '                     .primary_supplier_rec = rowPackingList.Primary_Supplier
    '                     .primary_transaksi = -1
    '                     .promise_to_deliver = rowPackingList.Estimated_Arrival_Date
    '                     .qty = rowPackingList.Total_Qty
    '                     .remarks = ""
    '                     .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.GIT)
    '                     .identifier = rowPackingList.identifier
    '                 End With
    '                 vTBLItemPOB.Rows.Add(rowItemPOB)
    '             Next
    '         End If

    '         'sob
    '         DTASOItem.Fill(TBLSOItem)
    '         If TBLSOItem.Rows.Count > 0 Then
    '             PrimaryPOMax = 0
    '             PrimaryPOMax = Double.MinValue
    '             For Each rowSOItem In TBLSOItem.Rows
    '                 If rowSOItem.Primary_Transaksi > PrimaryPOMax Then
    '                     PrimaryPOMax = rowSOItem.Primary_Transaksi
    '                 End If
    '                 If PrimaryPOMin > rowSOItem.Primary_Transaksi Then
    '                     PrimaryPOMin = rowSOItem.Primary_Transaksi
    '                 End If
    '             Next
    '             DTASOItemDelivery.FillBy_PrimarySOMaxMin(TBLSOItemDelivery, PrimaryPOMin, PrimaryPOMax)

    '             TBLSOItemDelivery.DefaultView.Sort = "[delivery date] DESC"
    '             For Each rowSOItem In TBLSOItem.Rows
    '                 TBLSOItemDelivery.DefaultView.RowFilter = "identifier = '" + rowSOItem.identifier + "'"

    '                 If TBLSOItemDelivery.DefaultView.Count = 0 Then
    '                     'tambahkan ke dalam item so sesuai sob
    '                     rowItemSO_SF = vTBLItemSOB.Newtbltrans_shipping_forecast_itemdeliveryRow
    '                     With rowItemSO_SF
    '                         .create_do = 0
    '                         'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery
    '                         'dan 0 untuk qty available
    '                         .qty_available = 0
    '                         .qty_ready_stock = 0
    '                         .estimate_delivery = #1/1/2200#
    '                         .primary = Me.GetPrimaryMax_ItemSOB(vTBLItemSOB) + 1
    '                         .primary_code_no = rowSOItem.Primary_Product
    '                         .primary_cust = rowSOItem.Primary_Customer
    '                         .primary_cust_rec = rowSOItem.Primary_Customer
    '                         .primary_priority_level = PrimaryPriorityLevelDefault
    '                         .primary_so = rowSOItem.Primary_Transaksi
    '                         .primary_transaksi = -1
    '                         .primary_unit = rowSOItem.Primary_Units
    '                         'tidak ada delivery control -> set tanggal ke 1 jan 2200
    '                         .promise_to_deliver = #1/1/2200#
    '                         .qty_sob = rowSOItem.SOB
    '                         .remark = ""
    '                         .so_created_by = rowSOItem.Created_By
    '                         .so_date_rec = rowSOItem.Tanggal_Transaksi
    '                         .so_number_rec = rowSOItem.No_Transaksi
    '                         .urutan_priority_level = UrutanPriodityLevelDefault
    '                         .customer_name = rowSOItem.Customer_Name
    '                         .Product_Code = rowSOItem.Product_Code
    '                         .Description = rowSOItem.Description
    '                         .so_number = rowSOItem.No_Transaksi
    '                         .so_date = rowSOItem.Tanggal_Transaksi
    '                         .Cust_PO_No = rowSOItem.Cust_PO_No
    '                         ._Decimal = rowSOItem._Decimal
    '                         .Unit_Code = rowSOItem.Unit_Code
    '                         .Primary_Category = rowSOItem.Primary_Category
    '                         .Primary_Sub_Brand = rowSOItem.Primary_Sub_Brand
    '                         .Primary_Brand = rowSOItem.Primary_Brand
    '                         .primary_product_group = rowSOItem.Primary_Product_Group
    '                         .identifier = rowSOItem.identifier
    '                         .booked = 0
    '                         .no_booking = ""
    '                         .qty_booked = 0
    '                     End With
    '                     vTBLItemSOB.Rows.Add(rowItemSO_SF)
    '                 Else
    '                     'ambil item delivery control mulai dari yang paling besar promise to deliverynya
    '                     While rowSOItem.SOB > 0
    '                         rowSOItemDelivery = TBLSOItemDelivery.DefaultView(0).Row
    '                         rowItemSO_SF = vTBLItemSOB.Newtbltrans_shipping_forecast_itemdeliveryRow
    '                         With rowItemSO_SF
    '                             .create_do = 0
    '                             'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery dan 0 untuk qty availabel
    '                             .estimate_delivery = #1/1/2200#
    '                             .qty_available = 0
    '                             .qty_ready_stock = 0
    '                             .primary = Me.GetPrimaryMax_ItemSOB(vTBLItemSOB) + 1
    '                             .primary_code_no = rowSOItem.Primary_Product
    '                             .primary_cust = rowSOItem.Primary_Customer
    '                             .primary_cust_rec = rowSOItem.Primary_Customer
    '                             .primary_priority_level = rowSOItemDelivery.delivery_priority
    '                             .urutan_priority_level = rowSOItemDelivery.urutan
    '                             .primary_so = rowSOItem.Primary_Transaksi
    '                             .primary_transaksi = -1
    '                             .primary_unit = rowSOItem.Primary_Units
    '                             'tidak ada delivery control -> set tanggal ke 1 jan 2200
    '                             .promise_to_deliver = rowSOItemDelivery.delivery_date
    '                             If rowSOItem.SOB > rowSOItemDelivery.delivery_qty Then
    '                                 .qty_sob = rowSOItemDelivery.delivery_qty
    '                                 rowSOItem.SOB -= rowSOItemDelivery.delivery_qty
    '                                 rowSOItemDelivery.delivery_qty = 0
    '                             Else
    '                                 .qty_sob = rowSOItem.SOB
    '                                 rowSOItem.SOB = 0
    '                             End If
    '                             .remark = ""
    '                             .so_created_by = rowSOItem.Created_By
    '                             .so_date_rec = rowSOItem.Tanggal_Transaksi
    '                             .so_number_rec = rowSOItem.No_Transaksi
    '                             .customer_name = rowSOItem.Customer_Name
    '                             .Product_Code = rowSOItem.Product_Code
    '                             .Description = rowSOItem.Description
    '                             .so_number = rowSOItem.No_Transaksi
    '                             .so_date = rowSOItem.Tanggal_Transaksi
    '                             .Cust_PO_No = rowSOItem.Cust_PO_No
    '                             ._Decimal = rowSOItem._Decimal
    '                             .Unit_Code = rowSOItem.Unit_Code
    '                             .Primary_Category = rowSOItem.Primary_Category
    '                             .Primary_Sub_Brand = rowSOItem.Primary_Sub_Brand
    '                             .Primary_Brand = rowSOItem.Primary_Brand
    '                             .primary_product_group = rowSOItem.Primary_Product_Group
    '                             .identifier = rowSOItem.identifier
    '                             .booked = 0
    '                             .no_booking = ""
    '                             .qty_booked = 0
    '                             If rowSOItemDelivery.delivery_qty = 0 Then
    '                                 TBLSOItemDelivery.Rows.Remove(rowSOItemDelivery)
    '                             End If
    '                         End With
    '                         vTBLItemSOB.Rows.Add(rowItemSO_SF)
    '                     End While
    '                 End If
    '             Next
    '         End If

    '         'tandain item sob yang ada inventory bookingnya
    '         DTABooking.FillBy_PrimaryCodeNo(TBLBooking, vPrimaryProduct)
    '         For Each rowBooking In TBLBooking.Rows
    '             vTBLItemSOB.DefaultView.RowFilter = "[primary code no] = " + rowBooking.Primary_Product.ToString + " AND [primary so] = " + rowBooking.Primary_SO.ToString
    '             vTBLItemSOB.DefaultView.Sort = "[promise to deliver] ASC, [urutan priority level] ASC, [so date rec] ASC, [primary so] ASC"
    '             If vTBLItemSOB.DefaultView.Count > 0 Then
    '                 IndexAcuan = 0
    '                 While IndexAcuan < vTBLItemSOB.DefaultView.Count And rowBooking.Qty > 0
    '                     rowItemSO_SF = vTBLItemSOB.DefaultView(IndexAcuan).Row

    '                     If rowItemSO_SF.qty_sob - rowItemSO_SF.qty_booked >= rowBooking.Qty Then
    '                         rowItemSO_SF.qty_booked += rowBooking.Qty
    '                         rowBooking.Qty = 0
    '                     Else
    '                         rowBooking.Qty -= (rowItemSO_SF.qty_sob - rowItemSO_SF.qty_booked)
    '                         rowItemSO_SF.qty_booked = rowItemSO_SF.qty_sob
    '                         IndexAcuan += 1
    '                     End If

    '                     rowItemSO_SF.booked = 1
    '                     If rowItemSO_SF.no_booking = "" Then
    '                         rowItemSO_SF.no_booking = rowBooking.No_Transaksi
    '                     Else
    '                         If rowItemSO_SF.no_booking.Contains(rowBooking.No_Transaksi) = False Then
    '                             rowItemSO_SF.no_booking = rowItemSO_SF.no_booking + ", " + rowBooking.No_Transaksi
    '                         End If
    '                     End If
    '                 End While
    '             End If
    '         Next

    '         'menghitung perkiraan kapan pengiriman selanjutnya
    '         For Each rowItemPOB In vTBLItemPOB.Rows
    '             'set parameter item pob yang akan digunakan pada perhitungan delivery date
    '             rowItemPOB.available_date = rowItemPOB.promise_to_deliver + TimeSpan.FromDays(rowItemPOB.delivery_delay)
    '             rowItemPOB.qty_used = 0
    '         Next
    '         'set parameter item sob yang akan digunakan pada perhitungan delivery date
    '         For Each rowItemSO_SF In vTBLItemSOB.Rows
    '             rowItemSO_SF.qty_available = 0
    '             rowItemSO_SF.estimate_delivery = #1/1/2200#
    '             rowItemSO_SF.qty_ready_stock = 0
    '         Next
    '         'set parameter item stock untuk perhitungan ulang
    '         For Each rowStock In vTBLItemStock.Rows
    '             rowStock.qty_used = 0
    '         Next

    '         'filter + sort
    '         vTBLItemPOB.DefaultView.Sort = "[available date] ASC, [promise to deliver] ASC, [po date rec] ASC, [primary po] ASC"
    '         vTBLItemSOB.DefaultView.Sort = "[booked] DESC, [promise to deliver] ASC, [urutan priority level] ASC, [so date rec] ASC, [primary so] ASC"

    '         'kembalikan parameter ke nilai awal
    '         IndexItemDelivery = 0
    '         QtyPOReady = 0
    '         QtyReadyStock = 0

    '         'pasangkan dengan yang ready stock dahulu
    '         If vTBLItemStock.DefaultView.Count > 0 Then
    '             rowStock = vTBLItemStock.DefaultView(0).Row

    '             While rowStock.qty - rowStock.qty_used > 0 And vTBLItemSOB.DefaultView.Count > IndexItemDelivery
    '                 rowItemSO_SF = vTBLItemSOB.DefaultView(IndexItemDelivery).Row
    '                 If rowStock.qty - rowStock.qty_used >= rowItemSO_SF.qty_sob Then
    '                     rowStock.qty_used += rowItemSO_SF.qty_sob
    '                     'karena ready stock, maka qty availabel = qty sob dan qty ready stock = qty sob
    '                     rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
    '                     rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
    '                     'karena ready stock, maka tandai estimate delivery menjadi ASAP -> 1 jan 1950
    '                     rowItemSO_SF.estimate_delivery = #1/1/1950#
    '                     'naikan index item delivery, cek item delivery selanjutnya
    '                     IndexItemDelivery += 1
    '                 Else
    '                     QtyPOReady = rowStock.qty - rowStock.qty_used
    '                     rowStock.qty_used = rowStock.qty
    '                     QtyReadyStock = QtyPOReady
    '                     Exit While
    '                 End If
    '             End While
    '         End If

    '         'pasangkan dengan item po
    '         If vTBLItemPOB.DefaultView.Count > 0 Then
    '             IndexItemPOB = 0
    '             While IndexItemPOB < vTBLItemPOB.DefaultView.Count And IndexItemDelivery < vTBLItemSOB.DefaultView.Count
    '                 rowItemPOB = vTBLItemPOB.DefaultView(IndexItemPOB).Row
    '                 rowItemSO_SF = vTBLItemSOB.DefaultView(IndexItemDelivery).Row
    '                 If QtyPOReady + rowItemPOB.qty - rowItemPOB.qty_used >= rowItemSO_SF.qty_sob Then
    '                     'jika item yang tersedia lebih besar dari item delivery, maka 
    '                     'set delivery date + qty available delivery + pindah ke baris item delivery selanjutnya
    '                     rowItemSO_SF.estimate_delivery = rowItemPOB.available_date
    '                     rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
    '                     'majukan index item delivery
    '                     IndexItemDelivery += 1

    '                     'kurangkan qty po ready
    '                     Dummy = rowItemSO_SF.qty_sob
    '                     If QtyPOReady >= Dummy Then
    '                         QtyPOReady -= Dummy
    '                         Dummy = 0
    '                     Else
    '                         Dummy -= QtyPOReady
    '                         QtyPOReady = 0
    '                     End If

    '                     'tambahkan qty used pada item po, untuk menanda jumlah yang sudah digunakan
    '                     rowItemPOB.qty_used += Dummy

    '                     'jika semua qty sudah habis dipakai, maka pindah ke baris item pob selanjutnya
    '                     If rowItemPOB.qty <= rowItemPOB.qty_used Then
    '                         IndexItemPOB += 1
    '                     End If

    '                     'perhitungan qty ready stock 
    '                     If QtyReadyStock >= rowItemSO_SF.qty_sob Then
    '                         rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
    '                         QtyReadyStock -= rowItemSO_SF.qty_sob
    '                     Else
    '                         rowItemSO_SF.qty_ready_stock = QtyReadyStock
    '                         QtyReadyStock = 0
    '                     End If
    '                 Else
    '                     'jika item delivery lebih besar dari qty po ready + qty po,ada 2 kemungkinan:
    '                     If IndexItemPOB + 1 < vTBLItemPOB.DefaultView.Count Then
    '                         '1. tambahkan qty po dari baris pob selanjut, karena dari pob, maka qty ready stock tidak di tambahkan
    '                         QtyPOReady += rowItemPOB.qty - rowItemPOB.qty_used
    '                         rowItemPOB.qty_used = rowItemPOB.qty
    '                         IndexItemPOB += 1
    '                     Else
    '                         '2. tidak ada baris pob selanjutnya, maka qty availble hanya sejumlah yang ada,
    '                         'estimate delivery tidak dikehui kapan, tapi available qty ada
    '                         rowItemSO_SF.qty_available = rowItemPOB.qty - rowItemPOB.qty_used + QtyPOReady
    '                         rowItemSO_SF.estimate_delivery = #1/1/2200#
    '                         'po ready di nol kan karena sudah habis
    '                         QtyPOReady = 0
    '                         'perhitungan ready stock
    '                         If QtyReadyStock >= rowItemSO_SF.qty_sob Then
    '                             rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
    '                             QtyReadyStock -= rowItemSO_SF.qty_sob
    '                         Else
    '                             rowItemSO_SF.qty_ready_stock = QtyReadyStock
    '                             QtyReadyStock = 0
    '                         End If
    '                         'tambahkan baris dgv
    '                         'tambahan index item delivery, untuk sisanya
    '                         IndexItemDelivery += 1
    '                         'index item pob tetap dinaikan, untuk menghentikan looping perhitungan
    '                         IndexItemPOB += 1
    '                     End If
    '                 End If
    '             End While
    '         End If

    '         'jika masih ada sisa item pob yang belum di handle, maka tambahkan ke dgv
    '         IndexItemPOB += 1
    '         While IndexItemPOB < vTBLItemPOB.DefaultView.Count
    '             rowItemPOB = vTBLItemPOB.DefaultView(IndexItemPOB).Row
    '             IndexItemPOB += 1
    '         End While
    '         'jika masih ada sisa item sob yang belum di handle, maka tambahkan ke dgv
    '         While IndexItemDelivery < vTBLItemSOB.DefaultView.Count
    '             rowItemSO_SF = vTBLItemSOB.DefaultView(IndexItemDelivery).Row
    '             If QtyPOReady > 0 Then
    '                 rowItemSO_SF.qty_available += QtyPOReady
    '                 QtyPOReady = 0
    '             End If
    '             If QtyReadyStock > 0 Then
    '                 rowItemSO_SF.qty_ready_stock += QtyReadyStock
    '                 QtyReadyStock = 0
    '             End If
    '             IndexItemDelivery += 1
    '         End While
    '     Catch ex As Exception
    '         MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '     Finally
    '         vTBLItemPOB.Dispose()
    '         DTA_ProductQty.Dispose()
    '         TBL_ProductQty.Dispose()
    '         DTAPendingDO.Dispose()
    '         TBLPendingDO.Dispose()
    '         ObjSOItem = Nothing
    '         DTAPOItem.Dispose()
    '         TBLPOItem.Dispose()
    '         DTAItemDeliveryPO.Dispose()
    '         TBLItemDeliveryPO.Dispose()
    '         DTAPackingList.Dispose()
    '         TBLPackingList.Dispose()
    '         DTASOItem.Dispose()
    '         TBLSOItem.Dispose()
    '         DTASOItemDelivery.Dispose()
    '         TBLSOItemDelivery.Dispose()
    '         ObjPriorityLevel = Nothing
    '         DTABooking.Dispose()
    '         TBLBooking.Dispose()
    '     End Try
    ' End Sub

    Private Function GetPrimaryMax_ItemSOB(ByVal vTBLItemDelivery As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable) As Double
        Dim rowsob As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim Retval As Double = 0

        Try
            For Each rowsob In vTBLItemDelivery.Rows
                If Retval < rowsob.primary Then
                    Retval = rowsob.primary
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetPrimaryMax_ItemPOB(ByVal vTBLItemPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable) As Double
        Dim rowpob As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow
        Dim Retval As Double = 0

        Try
            For Each rowpob In vTBLItemPOB.Rows
                If Retval < rowpob.primary Then
                    Retval = rowpob.primary
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetPrimaryMax_Stock(ByVal vTBLItemStock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable) As Double
        Dim rowStock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockRow
        Dim Retval As Double = 0

        Try
            For Each rowStock In vTBLItemStock.Rows
                If Retval < rowStock.primary Then
                    Retval = rowStock.primary
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub HitungShippingForecast(ByVal vTBLItemDelivery As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable, _
    ByVal vTBLItemPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable, _
    ByVal vTBLItemStock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable, ByVal vPrimarySF As Double, _
    ByVal vBackColorAll As Color, Optional ByVal vPrimaryCodeNo As Double = -1, Optional ByVal ListPrimaryCodeNo1 As List(Of Double) = Nothing)
        Dim ObjPriorityLevel As New Core_Module.ClsMasterPriorityLevel
        Dim DTA_ProductQty As New DseTransaksiShippingForecastTableAdapters.tblmaster_productqtyTableAdapter
        Dim TBL_ProductQty As New DseTransaksiShippingForecast.tblmaster_productqtyDataTable
        Dim row_ProductQty As DseTransaksiShippingForecast.tblmaster_productqtyRow
        Dim rowStock As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockRow
        Dim rowItemPOB As DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobRow
        Dim CekOtorisasiPO As Boolean = False
        Dim DTAPOItem As New DseTransaksiShippingForecastTableAdapters.tbltrans_po_itemTableAdapter
        Dim TBLPOItem As New DseTransaksiShippingForecast.tbltrans_po_itemDataTable
        Dim rowPOItem As DseTransaksiShippingForecast.tbltrans_po_itemRow
        Dim TBL_POB As New DseTransaksiMaterialPlanning2.DataTable3DataTable
        Dim ListSudah As New List(Of Double)
        Dim PrimaryPOMax, PrimaryPOMin As Double
        Dim DTAItemDeliveryPO As New DseTransaksiShippingForecastTableAdapters.tbltrans_po_itemdeliveryTableAdapter
        Dim TBLItemDeliveryPO As New DseTransaksiShippingForecast.tbltrans_po_itemdeliveryDataTable
        Dim rowItemDeliveryPO As DseTransaksiShippingForecast.tbltrans_po_itemdeliveryRow
        Dim DTAPackingList As New DseTransaksiShippingForecastTableAdapters.tblpackinglist_itemTableAdapter
        Dim TBLPackingList As New DseTransaksiShippingForecast.tblpackinglist_itemDataTable
        Dim rowPackingList As DseTransaksiShippingForecast.tblpackinglist_itemRow
        Dim DTASOItem As New DseTransaksiShippingForecastTableAdapters.tbltrans_so_itemTableAdapter
        Dim TBLSOItem As New DseTransaksiShippingForecast.tbltrans_so_itemDataTable
        Dim rowSOItem As DseTransaksiShippingForecast.tbltrans_so_itemRow
        Dim DTASOItemDelivery As New DseTransaksiShippingForecastTableAdapters.tbltrans_so_itemdeliveryTableAdapter
        Dim TBLSOItemDelivery As New DseTransaksiShippingForecast.tbltrans_so_itemdeliveryDataTable
        Dim rowSOItemDelivery As DseTransaksiShippingForecast.tbltrans_so_itemdeliveryRow
        Dim rowItemSO_SF As DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim PrimaryPriorityLevelDefault As Double = ObjPriorityLevel.GetPrimaryDefault
        Dim UrutanPriodityLevelDefault As Integer = ObjPriorityLevel.GetUrutan(PrimaryPriorityLevelDefault)

        Dim ListPrimaryCodeNo As New List(Of Double)
        Dim DTAPendingDO As New DseTransaksiShippingForecastTableAdapters.tbltrans_do_itemTableAdapter
        Dim TBLPendingDO As New DseTransaksiShippingForecast.tbltrans_do_itemDataTable
        Dim rowPendingDO As DseTransaksiShippingForecast.tbltrans_do_itemRow

        Dim DTABooking As New DseTransaksiShippingForecastTableAdapters.tbltrans_booking_invitemTableAdapter
        Dim TBLBooking As New DseTransaksiShippingForecast.tbltrans_booking_invitemDataTable
        Dim rowBooking As DseTransaksiShippingForecast.tbltrans_booking_invitemRow
        Dim IndexAcuan As Integer = 0
        Dim QtyPOReady As Decimal
        Dim QtyReadyStock As Decimal
        Dim QtyReadyStock1 As Decimal
        Dim IndexItemDelivery As Integer = 0
        Dim IndexItemPOB As Integer = 0
        Dim Dummy As Decimal = 0
        Dim FrmCek As New FrmTrans_MaterialPlanningDashboard_DetailPO

        Dim DTAPurchReqItem As New DseTransaksiShippingForecastTableAdapters.tbltrans_purch_req_itemTableAdapter
        Dim TBLPurchReqItem As New DseTransaksiShippingForecast.tbltrans_purch_req_itemDataTable
        Dim rowPurchReqItem As DseTransaksiShippingForecast.tbltrans_purch_req_itemRow
        Dim DTAPurchReqItemDelivery As New DseTransaksiShippingForecastTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
        Dim TBLPurchReqItemDelivery As New DseTransaksiShippingForecast.tbltrans_purchase_req_itemdeliveryDataTable
        Dim rowPurchReqItemDelivery As DseTransaksiShippingForecast.tbltrans_purchase_req_itemdeliveryRow
        Dim sdf As Double = 0
        Dim Lanjut As Boolean = False

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DTA_ProductQty.Connection.ConnectionString = Core_Module.GetConstring
            DTAPOItem.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTAPendingDO.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTAItemDeliveryPO.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTAPackingList.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTASOItemDelivery.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTABooking.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTAPurchReqItem.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString
            DTAPurchReqItemDelivery.Connection.ConnectionString = DTA_ProductQty.Connection.ConnectionString

            'kosongkan data awal
            vTBLItemDelivery.Rows.Clear()
            vTBLItemPOB.Rows.Clear()
            vTBLItemStock.Rows.Clear()

            'load data stock
            If vPrimaryCodeNo = -1 Then
                DTA_ProductQty.Fill(TBL_ProductQty)
                'If IsNothing(ListPrimaryCodeNo1) = False Then
                '    For gh As Integer = TBL_ProductQty.Rows.Count - 1 To 0 Step -1
                '        row_ProductQty = TBL_ProductQty.Rows(gh)
                '        If ListPrimaryCodeNo1.Contains(row_ProductQty.Primary_Product) = False Then
                '            TBL_ProductQty.Rows.RemoveAt(gh)
                '        End If
                '    Next
                'End If
            Else
                DTA_ProductQty.FillBy_PrimaryProduct(TBL_ProductQty, vPrimaryCodeNo)
            End If
            For Each row_ProductQty In TBL_ProductQty.Rows
                rowStock = vTBLItemStock.Newtbltrans_shipping_forecast_itemstockRow
                With rowStock
                    .primary = Me.GetPrimaryMax_Stock(vTBLItemStock) + 1
                    .primary_code_no = row_ProductQty.Primary_Product
                    .primary_transaksi = vPrimarySF
                    .primary_unit = row_ProductQty.Primary_Unit
                    .qty = row_ProductQty.Qty
                End With
                vTBLItemStock.Rows.Add(rowStock)
            Next
            'kurangkan data stock dengan pending do
            If vPrimaryCodeNo = -1 Then
                DTAPendingDO.Fill(TBLPendingDO)
                'If IsNothing(ListPrimaryCodeNo1) = False Then
                '    For gh As Integer = TBLPendingDO.Rows.Count - 1 To 0 Step -1
                '        rowPendingDO = TBLPendingDO.Rows(gh)
                '        If ListPrimaryCodeNo1.Contains(rowPendingDO.Primary_Product) = False Then
                '            TBLPendingDO.Rows.RemoveAt(gh)
                '        End If
                '    Next
                'End If
            Else
                DTAPendingDO.FillBy_PrimaryProduct(TBLPendingDO, vPrimaryCodeNo)
            End If
            For Each rowPendingDO In TBLPendingDO.Rows
                vTBLItemStock.DefaultView.RowFilter = "[primary code no] = " + rowPendingDO.Primary_Product.ToString
                If vTBLItemStock.DefaultView.Count > 0 Then
                    rowStock = vTBLItemStock.DefaultView(0).Row
                    rowStock.qty -= rowPendingDO.Total_Qty
                    If rowStock.qty <= 0 Then
                        vTBLItemStock.Rows.Remove(rowStock)
                    End If
                End If
            Next

            'load data pob
            If vPrimaryCodeNo = -1 Then
                DTAPOItem.Fill(TBLPOItem)
                'If IsNothing(ListPrimaryCodeNo1) = False Then
                '    For gh As Integer = TBLPOItem.Rows.Count - 1 To 0 Step -1
                '        rowPOItem = TBLPOItem.Rows(gh)
                '        If ListPrimaryCodeNo1.Contains(rowPOItem.Primary_Product) = False Then
                '            TBLPOItem.Rows.RemoveAt(gh)
                '        End If
                '    Next
                'End If
            Else
                DTAPOItem.FillBy_PrimaryProduct(TBLPOItem, vPrimaryCodeNo)
            End If
            'kosongkan item po yang belum diotorisasi
            IndexItemPOB = 0
            While IndexItemPOB < TBLPOItem.Rows.Count
                rowPOItem = TBLPOItem.Rows(IndexItemPOB)
                If rowPOItem.Authorization_Level <> rowPOItem.Authorization_Level_Needed Then
                    TBLPOItem.Rows.Remove(rowPOItem)
                Else
                    IndexItemPOB += 1
                End If
            End While

            'pindahkan ke dalam item pob
            'cari primary po max dan min
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If TBLPOItem.Rows.Count > 0 Then
                PrimaryPOMin = Double.MaxValue
                PrimaryPOMax = 0
                For Each rowPOItem In TBLPOItem.Rows
                    If PrimaryPOMax < rowPOItem.Primary_Transaksi Then
                        PrimaryPOMax = rowPOItem.Primary_Transaksi
                    End If
                    If PrimaryPOMin > rowPOItem.Primary_Transaksi Then
                        PrimaryPOMin = rowPOItem.Primary_Transaksi
                    End If
                Next
                'load item delivery
                If vPrimaryCodeNo = -1 Then
                    DTAItemDeliveryPO.Fill(TBLItemDeliveryPO, PrimaryPOMax, PrimaryPOMin)
                    'If IsNothing(ListPrimaryCodeNo1) = False Then
                    '    For gh As Integer = TBLItemDeliveryPO.Rows.Count - 1 To 0 Step -1
                    '        rowItemDeliveryPO = TBLItemDeliveryPO.Rows(gh)
                    '        If ListPrimaryCodeNo1.Contains(rowItemDeliveryPO.Primary_Product) = False Then
                    '            TBLItemDeliveryPO.Rows.RemoveAt(gh)
                    '        End If
                    '    Next
                    'End If
                Else
                    DTAItemDeliveryPO.FillBy_PrimaryProduct(TBLItemDeliveryPO, PrimaryPOMax, PrimaryPOMin, vPrimaryCodeNo)
                End If
                TBLItemDeliveryPO.DefaultView.Sort = "[delivery date] DESC"
                'isikan ke dalam tabel item pob
                For Each rowPOItem In TBLPOItem.Rows
                    TBLItemDeliveryPO.DefaultView.RowFilter = "identifier = '" + rowPOItem.identifier + "'"
                    If TBLItemDeliveryPO.DefaultView.Count = 0 Then
                        'tidak ada delivery control -> set tanggal ke 1 jan 2200
                        rowItemPOB = vTBLItemPOB.Newtbltrans_shipping_forecast_item_pobRow
                        With rowItemPOB
                            .delivery_delay = 0
                            .packing_list_date = #1/1/1950#
                            .packing_list_no_rec = ""
                            .po_date_rec = rowPOItem.PO_Date
                            .po_number_rec = rowPOItem.PO_Number
                            .primary = GetPrimaryMax_ItemPOB(vTBLItemPOB) + 1
                            .primary_code_no = rowPOItem.Primary_Product
                            .primary_packing_list = -1
                            .primary_po = rowPOItem.Primary_Transaksi
                            .primary_supplier = rowPOItem.Primary_Supplier
                            .primary_supplier_rec = rowPOItem.Primary_Supplier
                            .primary_transaksi = vPrimarySF
                            .promise_to_deliver = #1/1/2200#
                            .qty = rowPOItem.POB
                            .remarks = ""
                            .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.POB)
                            .identifier = rowPOItem.identifier
                        End With
                        vTBLItemPOB.Rows.Add(rowItemPOB)
                    Else
                        'ambil item delivery control mulai dari yang paling besar promise to deliverynya
                        While rowPOItem.POB > 0
                            rowItemDeliveryPO = TBLItemDeliveryPO.DefaultView(0).Row
                            rowItemPOB = vTBLItemPOB.Newtbltrans_shipping_forecast_item_pobRow
                            With rowItemPOB
                                .delivery_delay = rowItemDeliveryPO.delay
                                .packing_list_date = #1/1/1950#
                                .packing_list_no_rec = ""
                                .po_date_rec = rowPOItem.PO_Date
                                .po_number_rec = rowPOItem.PO_Number
                                .primary = GetPrimaryMax_ItemPOB(vTBLItemPOB) + 1
                                .primary_code_no = rowPOItem.Primary_Product
                                .primary_packing_list = -1
                                .primary_po = rowPOItem.Primary_Transaksi
                                .primary_supplier = rowPOItem.Primary_Supplier
                                .primary_supplier_rec = rowPOItem.Primary_Supplier
                                .primary_transaksi = vPrimarySF
                                .promise_to_deliver = rowItemDeliveryPO.delivery_date
                                If rowPOItem.POB > rowItemDeliveryPO.delivery_qty Then
                                    .qty = rowItemDeliveryPO.delivery_qty
                                    rowPOItem.POB -= rowItemDeliveryPO.delivery_qty
                                    rowItemDeliveryPO.delivery_qty = 0
                                Else
                                    .qty = rowPOItem.POB
                                    rowPOItem.POB = 0
                                End If
                                .remarks = ""
                                .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.POB)
                                .identifier = rowPOItem.identifier
                            End With
                            vTBLItemPOB.Rows.Add(rowItemPOB)

                            'persiapan untuk looping selanjutnya
                            If rowItemDeliveryPO.delivery_qty = 0 Then
                                TBLItemDeliveryPO.Rows.Remove(rowItemDeliveryPO)
                            End If
                        End While
                    End If
                Next
            End If

            'load data git
            If vPrimaryCodeNo = -1 Then
                DTAPackingList.Fill(TBLPackingList)
                'If IsNothing(ListPrimaryCodeNo1) = False Then
                '    For gh As Integer = TBLPackingList.Rows.Count - 1 To 0 Step -1
                '        rowPackingList = TBLPackingList.Rows(gh)
                '        If ListPrimaryCodeNo1.Contains(rowPackingList.Primary_Product) = False Then
                '            TBLPackingList.Rows.RemoveAt(gh)
                '        End If
                '    Next
                'End If
            Else
                DTAPackingList.FillBy_PrimaryProduct(TBLPackingList, vPrimaryCodeNo)
            End If

            'isikan data delay
            If TBLPackingList.Rows.Count > 0 Then
                PrimaryPOMax = 0
                PrimaryPOMin = Double.MaxValue
                For Each rowPackingList In TBLPackingList.Rows
                    If PrimaryPOMax < rowPackingList.Primary_PO Then
                        PrimaryPOMax = rowPackingList.Primary_PO
                    End If
                    If PrimaryPOMin > rowPackingList.Primary_PO Then
                        PrimaryPOMin = rowPackingList.Primary_PO
                    End If
                Next
                If vPrimaryCodeNo = -1 Then
                    DTAItemDeliveryPO.Fill(TBLItemDeliveryPO, PrimaryPOMax, PrimaryPOMin)
                Else
                    DTAItemDeliveryPO.FillBy_PrimaryProduct(TBLItemDeliveryPO, PrimaryPOMax, PrimaryPOMin, vPrimaryCodeNo)
                End If
                TBLPOItem.Rows.Clear()
                For Each rowPackingList In TBLPackingList.Rows
                    TBLItemDeliveryPO.DefaultView.RowFilter = "[primary transaksi] = " + rowPackingList.Primary_PO.ToString + " AND [Primary Product] = " + rowPackingList.Primary_Product.ToString
                    If TBLItemDeliveryPO.DefaultView.Count > 0 Then
                        rowItemDeliveryPO = TBLItemDeliveryPO.DefaultView(0).Row
                        rowPackingList.delay = rowItemDeliveryPO.delay
                        rowPackingList.identifier = rowItemDeliveryPO.identifier
                    Else
                        rowPackingList.delay = 0
                        If TBLPOItem.Rows.Count = 0 Then
                            DTAPOItem.FillBy_PrimaryPOMax_Min(TBLPOItem, PrimaryPOMax, PrimaryPOMin)
                        End If
                        TBLPOItem.DefaultView.RowFilter = "[primary transaksi] = " + rowPackingList.Primary_PO.ToString + " AND [Primary Product] = " + rowPackingList.Primary_Product.ToString
                        rowPOItem = TBLPOItem.DefaultView(0).Row
                        rowPackingList.identifier = rowPOItem.identifier
                    End If
                Next
                'pindahkan ke dalam item packinglist
                For Each rowPackingList In TBLPackingList.Rows
                    rowItemPOB = vTBLItemPOB.Newtbltrans_shipping_forecast_item_pobRow
                    With rowItemPOB
                        .delivery_delay = rowPackingList.delay
                        .packing_list_date = rowPackingList.Tanggal_Transaksi
                        .packing_list_no_rec = rowPackingList.No_Transaksi
                        'untuk po date rec, diisikan dengan tanggal packing list, untuk pengurutan data ketika akan menentukan item delivery
                        .po_date_rec = #1/1/1950#
                        'untuk po number, diisikan dengan supplier packing list number, untuk record jika suatu saat diperlukan
                        .po_number_rec = rowPackingList.Supplier_Packing_List_No
                        .primary = GetPrimaryMax_ItemPOB(vTBLItemPOB) + 1
                        .primary_code_no = rowPackingList.Primary_Product
                        .primary_packing_list = rowPackingList.Primary_Packing_List
                        .primary_po = rowPackingList.Primary_PO
                        .primary_supplier = rowPackingList.Primary_Supplier
                        .primary_supplier_rec = rowPackingList.Primary_Supplier
                        .primary_transaksi = vPrimarySF
                        .promise_to_deliver = rowPackingList.Estimated_Arrival_Date
                        .qty = rowPackingList.Total_Qty
                        .remarks = ""
                        .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemPOB.GIT)
                        .identifier = rowPackingList.identifier
                    End With
                    vTBLItemPOB.Rows.Add(rowItemPOB)
                Next
            End If

            'load data sob
            If vPrimaryCodeNo = -1 Then
                DTASOItem.Fill(TBLSOItem)
                'If IsNothing(ListPrimaryCodeNo1) = False Then
                '    For gh As Integer = TBLSOItem.Rows.Count - 1 To 0 Step -1
                '        rowSOItem = TBLSOItem.Rows(gh)
                '        If ListPrimaryCodeNo1.Contains(rowSOItem.Primary_Product) = False Then
                '            TBLSOItem.Rows.RemoveAt(gh)
                '        End If
                '    Next
                'End If
            Else
                DTASOItem.FillBy_PrimaryProduct(TBLSOItem, vPrimaryCodeNo)
            End If
            If TBLSOItem.Rows.Count > 0 Then
                PrimaryPOMax = 0
                PrimaryPOMax = Double.MinValue
                For Each rowSOItem In TBLSOItem.Rows
                    If rowSOItem.Primary_Transaksi > PrimaryPOMax Then
                        PrimaryPOMax = rowSOItem.Primary_Transaksi
                    End If
                    If PrimaryPOMin > rowSOItem.Primary_Transaksi Then
                        PrimaryPOMin = rowSOItem.Primary_Transaksi
                    End If
                Next
                DTASOItemDelivery.FillBy_PrimarySOMaxMin(TBLSOItemDelivery, PrimaryPOMin, PrimaryPOMax)
                'If IsNothing(ListPrimaryCodeNo1) = False Then
                '    For gh As Integer = TBLSOItemDelivery.Rows.Count - 1 To 0 Step -1
                '        rowSOItemDelivery = TBLSOItemDelivery.Rows(gh)
                '        'if ListPrimaryCodeNo1 .Contains (rowSOItemDelivery .prim
                '    Next
                'End If

                TBLSOItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                For Each rowSOItem In TBLSOItem.Rows
                    TBLSOItemDelivery.DefaultView.RowFilter = "identifier = '" + rowSOItem.identifier + "'"

                    If TBLSOItemDelivery.DefaultView.Count = 0 Then
                        'tambahkan ke dalam item so sesuai sob
                        rowItemSO_SF = vTBLItemDelivery.Newtbltrans_shipping_forecast_itemdeliveryRow
                        With rowItemSO_SF
                            .create_do = 0
                            'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery
                            'dan 0 untuk qty available
                            .qty_available = 0
                            .qty_ready_stock = 0
                            .estimate_delivery = #1/1/2200#
                            .primary = Me.GetPrimaryMax_ItemSOB(vTBLItemDelivery) + 1
                            .primary_code_no = rowSOItem.Primary_Product
                            .primary_cust = rowSOItem.Primary_Customer
                            .primary_cust_rec = rowSOItem.Primary_Customer
                            .primary_priority_level = PrimaryPriorityLevelDefault
                            .primary_so = rowSOItem.Primary_Transaksi
                            .primary_transaksi = vPrimarySF
                            .primary_unit = rowSOItem.Primary_Units
                            'tidak ada delivery control -> set tanggal ke 1 jan 2200
                            .promise_to_deliver = #1/1/2200#
                            .qty_sob = rowSOItem.SOB
                            .remark = ""
                            .so_created_by = rowSOItem.Created_By
                            .so_date_rec = rowSOItem.Tanggal_Transaksi
                            .so_number_rec = rowSOItem.No_Transaksi
                            .urutan_priority_level = UrutanPriodityLevelDefault
                            .customer_name = rowSOItem.Customer_Name
                            .Product_Code = rowSOItem.Product_Code
                            .Description = rowSOItem.Description
                            .so_number = rowSOItem.No_Transaksi
                            .so_date = rowSOItem.Tanggal_Transaksi
                            .Cust_PO_No = rowSOItem.Cust_PO_No
                            ._Decimal = rowSOItem._Decimal
                            .Unit_Code = rowSOItem.Unit_Code
                            .Primary_Category = rowSOItem.Primary_Category
                            .Primary_Sub_Brand = rowSOItem.Primary_Sub_Brand
                            .Primary_Brand = rowSOItem.Primary_Brand
                            .primary_product_group = rowSOItem.Primary_Product_Group
                            .identifier = rowSOItem.identifier
                            .booked = 0
                            .no_booking = ""
                            .qty_booked = 0
                            .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB)
                            .tipe_product = rowSOItem.Type_Product
                        End With
                        vTBLItemDelivery.Rows.Add(rowItemSO_SF)
                    Else
                        'ambil item delivery control mulai dari yang paling besar promise to deliverynya
                        While rowSOItem.SOB > 0
                            rowSOItemDelivery = TBLSOItemDelivery.DefaultView(0).Row
                            rowItemSO_SF = vTBLItemDelivery.Newtbltrans_shipping_forecast_itemdeliveryRow
                            With rowItemSO_SF
                                .create_do = 0
                                'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery dan 0 untuk qty availabel
                                .estimate_delivery = #1/1/2200#
                                .qty_available = 0
                                .qty_ready_stock = 0
                                .primary = Me.GetPrimaryMax_ItemSOB(vTBLItemDelivery) + 1
                                .primary_code_no = rowSOItem.Primary_Product
                                .primary_cust = rowSOItem.Primary_Customer
                                .primary_cust_rec = rowSOItem.Primary_Customer
                                .primary_priority_level = rowSOItemDelivery.delivery_priority
                                .urutan_priority_level = rowSOItemDelivery.urutan
                                .primary_so = rowSOItem.Primary_Transaksi
                                .primary_transaksi = vPrimarySF
                                .primary_unit = rowSOItem.Primary_Units
                                'tidak ada delivery control -> set tanggal ke 1 jan 2200
                                .promise_to_deliver = rowSOItemDelivery.delivery_date
                                If rowSOItem.SOB > rowSOItemDelivery.delivery_qty Then
                                    .qty_sob = rowSOItemDelivery.delivery_qty
                                    rowSOItem.SOB -= rowSOItemDelivery.delivery_qty
                                    rowSOItemDelivery.delivery_qty = 0
                                Else
                                    .qty_sob = rowSOItem.SOB
                                    rowSOItem.SOB = 0
                                End If
                                .remark = ""
                                .so_created_by = rowSOItem.Created_By
                                .so_date_rec = rowSOItem.Tanggal_Transaksi
                                .so_number_rec = rowSOItem.No_Transaksi
                                .customer_name = rowSOItem.Customer_Name
                                .Product_Code = rowSOItem.Product_Code
                                .Description = rowSOItem.Description
                                .so_number = rowSOItem.No_Transaksi
                                .so_date = rowSOItem.Tanggal_Transaksi
                                .Cust_PO_No = rowSOItem.Cust_PO_No
                                ._Decimal = rowSOItem._Decimal
                                .Unit_Code = rowSOItem.Unit_Code
                                .Primary_Category = rowSOItem.Primary_Category
                                .Primary_Sub_Brand = rowSOItem.Primary_Sub_Brand
                                .Primary_Brand = rowSOItem.Primary_Brand
                                .primary_product_group = rowSOItem.Primary_Product_Group
                                .identifier = rowSOItem.identifier
                                .booked = 0
                                .no_booking = ""
                                .qty_booked = 0
                                .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.SOB)
                                .tipe_product = rowSOItem.Type_Product
                                If rowSOItemDelivery.delivery_qty = 0 Then
                                    TBLSOItemDelivery.Rows.Remove(rowSOItemDelivery)
                                End If
                            End With
                            vTBLItemDelivery.Rows.Add(rowItemSO_SF)
                        End While
                    End If
                Next
            End If

            'tandain item sob yang ada inventory bookingnya
            If vPrimaryCodeNo = -1 Then
                DTABooking.Fill(TBLBooking)
            Else
                DTABooking.FillBy_PrimaryCodeNo(TBLBooking, vPrimaryCodeNo)
            End If
            For Each rowBooking In TBLBooking.Rows
                vTBLItemDelivery.DefaultView.RowFilter = "[primary code no] = " + rowBooking.Primary_Product.ToString + " AND [primary so] = " + rowBooking.Primary_SO.ToString
                vTBLItemDelivery.DefaultView.Sort = "[promise to deliver] ASC, [urutan priority level] ASC, [so date rec] ASC, [primary so] ASC"
                IndexAcuan = 0
                While IndexAcuan < vTBLItemDelivery.DefaultView.Count And rowBooking.Qty > 0
                    rowItemSO_SF = vTBLItemDelivery.DefaultView(IndexAcuan).Row

                    If rowItemSO_SF.qty_sob - rowItemSO_SF.qty_booked >= rowBooking.Qty Then
                        rowItemSO_SF.qty_booked += rowBooking.Qty
                        rowBooking.Qty = 0
                    Else
                        rowBooking.Qty -= (rowItemSO_SF.qty_sob - rowItemSO_SF.qty_booked)
                        rowItemSO_SF.qty_booked = rowItemSO_SF.qty_sob
                        IndexAcuan += 1
                    End If

                    rowItemSO_SF.booked = 1
                    If rowItemSO_SF.no_booking = "" Then
                        rowItemSO_SF.no_booking = rowBooking.No_Transaksi
                    Else
                        If rowItemSO_SF.no_booking.Contains(rowBooking.No_Transaksi) = False Then
                            rowItemSO_SF.no_booking = rowItemSO_SF.no_booking + ", " + rowBooking.No_Transaksi
                        End If
                    End If
                End While
            Next

            'tambahkan item sob dengan purchase request
            DTAPurchReqItem.Fill(TBLPurchReqItem)
            If TBLPurchReqItem.Rows.Count > 0 Then
                PrimaryPOMax = 0
                PrimaryPOMax = Double.MinValue
                For Each rowPurchReqItem In TBLPurchReqItem.Rows
                    If rowPurchReqItem.primary_transaksi > PrimaryPOMax Then
                        PrimaryPOMax = rowPurchReqItem.primary_transaksi
                    End If
                    If PrimaryPOMin > rowPurchReqItem.primary_transaksi Then
                        PrimaryPOMin = rowPurchReqItem.primary_transaksi
                    End If
                Next
                DTAPurchReqItemDelivery.Fill(TBLPurchReqItemDelivery, PrimaryPOMax, PrimaryPOMin)

                TBLPurchReqItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                For Each rowPurchReqItem In TBLPurchReqItem.Rows
                    TBLPurchReqItemDelivery.DefaultView.RowFilter = "identifier = '" + rowPurchReqItem.identifier + "'"

                    If TBLPurchReqItemDelivery.DefaultView.Count = 0 Then
                        'tambahkan ke dalam item sob sesuai Purc Req OB
                        rowItemSO_SF = vTBLItemDelivery.Newtbltrans_shipping_forecast_itemdeliveryRow
                        With rowItemSO_SF
                            .create_do = 0
                            'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery
                            'dan 0 untuk qty available
                            .qty_available = 0
                            .qty_ready_stock = 0
                            .estimate_delivery = #1/1/2200#
                            .primary = Me.GetPrimaryMax_ItemSOB(vTBLItemDelivery) + 1
                            .primary_code_no = rowPurchReqItem.primary_code_no
                            .primary_cust = rowPurchReqItem.primary_ref_party
                            .primary_cust_rec = rowPurchReqItem.primary_ref_party
                            .primary_priority_level = PrimaryPriorityLevelDefault
                            .primary_so = rowPurchReqItem.primary_transaksi
                            .primary_transaksi = vPrimarySF
                            .primary_unit = rowPurchReqItem.primary_unit
                            'tidak ada delivery control -> set tanggal ke 1 jan 2200
                            .promise_to_deliver = #1/1/2200#
                            .qty_sob = rowPurchReqItem.OB
                            .remark = rowPurchReqItem.remarks_item
                            .so_created_by = rowPurchReqItem.created_by
                            .so_date_rec = rowPurchReqItem.tanggal
                            .so_number_rec = rowPurchReqItem.no_transaksi
                            .urutan_priority_level = UrutanPriodityLevelDefault
                            .customer_name = rowPurchReqItem.Ref_Party_Desc
                            .Product_Code = rowPurchReqItem.Product_Code
                            .Description = rowPurchReqItem.Description
                            .so_number = rowPurchReqItem.no_transaksi
                            .so_date = rowPurchReqItem.tanggal
                            .Cust_PO_No = ""
                            ._Decimal = rowPurchReqItem._Decimal
                            .Unit_Code = rowPurchReqItem.Unit_Code
                            .Primary_Category = rowPurchReqItem.Primary_Category
                            .Primary_Sub_Brand = rowPurchReqItem.Primary_Sub_Brand
                            .Primary_Brand = rowPurchReqItem.Primary_Brand
                            .primary_product_group = rowPurchReqItem.Primary_Product_Group
                            .identifier = rowPurchReqItem.identifier
                            .booked = 0
                            .no_booking = ""
                            .qty_booked = 0
                            .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req)
                            .tipe_product = rowPurchReqItem.Type
                        End With
                        vTBLItemDelivery.Rows.Add(rowItemSO_SF)
                    Else
                        'ambil item delivery control mulai dari yang paling besar promise to deliverynya
                        While rowPurchReqItem.OB > 0
                            rowPurchReqItemDelivery = TBLPurchReqItemDelivery.DefaultView(0).Row
                            rowItemSO_SF = vTBLItemDelivery.Newtbltrans_shipping_forecast_itemdeliveryRow
                            With rowItemSO_SF
                                .create_do = 0
                                'kolom ini dihitung kemudian, sementara diisikan 1 jan 2200 untuk estimate delivery dan 0 untuk qty availabel
                                .estimate_delivery = #1/1/2200#
                                .qty_available = 0
                                .qty_ready_stock = 0
                                .primary = Me.GetPrimaryMax_ItemSOB(vTBLItemDelivery) + 1
                                .primary_code_no = rowPurchReqItem.primary_code_no
                                .primary_cust = rowPurchReqItem.primary_ref_party
                                .primary_cust_rec = rowPurchReqItem.primary_ref_party
                                .primary_priority_level = rowPurchReqItemDelivery.delivery_priority
                                .urutan_priority_level = rowPurchReqItemDelivery.urutan
                                .primary_so = rowPurchReqItem.primary_transaksi
                                .primary_transaksi = vPrimarySF
                                .primary_unit = rowPurchReqItem.primary_unit
                                'tidak ada delivery control -> set tanggal ke 1 jan 2200
                                .promise_to_deliver = rowPurchReqItemDelivery.delivery_date
                                If rowPurchReqItem.OB > rowPurchReqItemDelivery.delivery_qty Then
                                    .qty_sob = rowPurchReqItemDelivery.delivery_qty
                                    rowPurchReqItem.OB -= rowPurchReqItemDelivery.delivery_qty
                                    rowPurchReqItemDelivery.delivery_qty = 0
                                Else
                                    .qty_sob = rowPurchReqItem.OB
                                    rowPurchReqItem.OB = 0
                                End If
                                .remark = rowPurchReqItem.remarks_item
                                .so_created_by = rowPurchReqItem.created_by
                                .so_date_rec = rowPurchReqItem.tanggal
                                .so_number_rec = rowPurchReqItem.no_transaksi
                                .customer_name = rowPurchReqItem.Ref_Party_Desc
                                .Product_Code = rowPurchReqItem.Product_Code
                                .Description = rowPurchReqItem.Description
                                .so_number = rowPurchReqItem.no_transaksi
                                .so_date = rowPurchReqItem.tanggal
                                .Cust_PO_No = ""
                                ._Decimal = rowPurchReqItem._Decimal
                                .Unit_Code = rowPurchReqItem.Unit_Code
                                .Primary_Category = rowPurchReqItem.Primary_Category
                                .Primary_Sub_Brand = rowPurchReqItem.Primary_Sub_Brand
                                .Primary_Brand = rowPurchReqItem.Primary_Brand
                                .primary_product_group = rowPurchReqItem.Primary_Product_Group
                                .identifier = rowPurchReqItem.identifier
                                .booked = 0
                                .no_booking = ""
                                .qty_booked = 0
                                .tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req)
                                .tipe_product = rowPurchReqItem.Type
                                If rowPurchReqItemDelivery.delivery_qty = 0 Then
                                    TBLPurchReqItemDelivery.Rows.Remove(rowPurchReqItemDelivery)
                                End If
                            End With
                            vTBLItemDelivery.Rows.Add(rowItemSO_SF)
                        End While
                    End If
                Next
            End If

            'pasangkan data sob dengan item stock dan pob, tentukan estimate delivery
            'delivery qty hanya bisa tidak sama dengan qty sob, bila barang tersedia sebagian dan belum ada order.
            'pasangkan terlebih dahulu yang ada inventory bookingnya
            'baru kemudia sisanya
            ListPrimaryCodeNo.Clear()
            For Each rowItemSO_SF In vTBLItemDelivery
                'If rowItemSO_SF.tipe_product = Core_Module.clsMasterProduct.EnumTipeProduct.Services Then
                '    'karena ready stock, maka qty availabel = qty sob
                '    rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
                '    rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                '    'karena ready stock, maka tandai estimate delivery menjadi ASAP -> 1 jan 1950
                '    rowItemSO_SF.estimate_delivery = #1/1/1950#
                'Else
                'untuk shipping forecast ini, service tidak dianggap ready, tapi diperlakukan sama dengan goods
                'kalo di shipping forecast yang biasa, 
                If ListPrimaryCodeNo.Contains(rowItemSO_SF.primary_code_no) = False Then
                    ListPrimaryCodeNo.Add(rowItemSO_SF.primary_code_no)
                End If
                rowItemSO_SF.primary_po_ref = ""
                'End If
                rowItemSO_SF.qty_available = 0
                rowItemSO_SF.qty_ready_stock = 0
                'karena ready stock, maka tandai estimate delivery menjadi ASAP -> 1 jan 1950
                rowItemSO_SF.estimate_delivery = #1/1/2200#
            Next
            For Each rowItemPOB In vTBLItemPOB.Rows
                'set parameter item pob yang akan digunakan pada perhitungan delivery date
                rowItemPOB.available_date = rowItemPOB.promise_to_deliver + TimeSpan.FromDays(rowItemPOB.delivery_delay)
                rowItemPOB.qty_used = 0
            Next
            For Each rowStock In vTBLItemStock.Rows
                rowStock.qty_used = 0
            Next

            For Each PrimaryCodeNo As Double In ListPrimaryCodeNo
                Lanjut = True
                If IsNothing(ListPrimaryCodeNo1) = False Then
                    If ListPrimaryCodeNo1.Count = 0 Then
                        Lanjut = True
                    Else
                        If ListPrimaryCodeNo1.Contains(PrimaryCodeNo) = False Then
                            Lanjut = False
                        Else
                            Lanjut = True
                        End If
                    End If
                End If

                If Lanjut = True Then
                    'filter
                    vTBLItemStock.DefaultView.RowFilter = "[primary code no] = " + PrimaryCodeNo.ToString
                    vTBLItemPOB.DefaultView.RowFilter = "[primary code no] = " + PrimaryCodeNo.ToString
                    vTBLItemPOB.DefaultView.Sort = "[available date] ASC, [promise to deliver] ASC, [po date rec] ASC, [primary po] ASC"
                    vTBLItemDelivery.DefaultView.RowFilter = "[primary code no] = " + PrimaryCodeNo.ToString
                    vTBLItemDelivery.DefaultView.Sort = "[booked] DESC, [promise to deliver] ASC, [urutan priority level] ASC, [so date rec] ASC, [primary so] ASC"

                    'kembalikan parameter ke nilai awal
                    IndexItemDelivery = 0
                    QtyPOReady = 0
                    QtyReadyStock = 0

                    'pasangkan dengan yang ready stock dahulu
                    'kalo tipenya prOB, maka tidak masuk sini, karena PrOB, hanya bisa di counter dengan POB atau GIT
                    If vTBLItemStock.DefaultView.Count > 0 Then
                        rowStock = vTBLItemStock.DefaultView(0).Row
                        While rowStock.qty - rowStock.qty_used > 0 And vTBLItemDelivery.DefaultView.Count > IndexItemDelivery
                            rowItemSO_SF = vTBLItemDelivery.DefaultView(IndexItemDelivery).Row
                            If rowItemSO_SF.tipe = CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req) Then
                                IndexItemDelivery += 1
                            Else
                                If rowStock.qty - rowStock.qty_used >= rowItemSO_SF.qty_sob Then
                                    rowStock.qty_used += rowItemSO_SF.qty_sob
                                    'karena ready stock, maka qty availabel = qty sob
                                    rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
                                    rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                                    'karena ready stock, maka tandai estimate delivery menjadi ASAP -> 1 jan 1950
                                    rowItemSO_SF.estimate_delivery = #1/1/1950#
                                    'naikan index item delivery, cek item delivery selanjutnya
                                    IndexItemDelivery += 1
                                Else
                                    QtyPOReady = rowStock.qty - rowStock.qty_used
                                    QtyReadyStock = QtyPOReady
                                    rowStock.qty_used = rowStock.qty
                                    Exit While
                                End If
                            End If
                        End While
                    End If

                    'pasangkan dengan item po
                    'IndexItemDelivery = 0 '-> ulang lagi dari yang paling atas, hanya saja untuk yang rowItemSO_SF.qty_available < rowItemSO_SF.qty_sob
                    If vTBLItemPOB.DefaultView.Count > 0 Then
                        IndexItemPOB = 0
                        While IndexItemPOB < vTBLItemPOB.DefaultView.Count And IndexItemDelivery < vTBLItemDelivery.DefaultView.Count
                            rowItemPOB = vTBLItemPOB.DefaultView(IndexItemPOB).Row
                            rowItemSO_SF = vTBLItemDelivery.DefaultView(IndexItemDelivery).Row
                            'jika pr-ob, maka tidak boleh meperhitungkan ready stock
                            If rowItemSO_SF.tipe <> CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req) Then
                                QtyReadyStock1 = QtyReadyStock
                            Else
                                QtyReadyStock1 = 0
                            End If
                            If QtyReadyStock1 + QtyPOReady + rowItemPOB.qty - rowItemPOB.qty_used >= rowItemSO_SF.qty_sob Then
                                'jika item yang tersedia lebih besar dari item delivery, maka 
                                'set delivery date + qty available delivery + pindah ke baris item delivery selanjutnya
                                rowItemSO_SF.estimate_delivery = rowItemPOB.available_date
                                rowItemSO_SF.qty_available = rowItemSO_SF.qty_sob
                                IndexItemDelivery += 1

                                'kurangkan qty on hand dan qty po ready
                                Dummy = rowItemSO_SF.qty_sob
                                If QtyReadyStock1 >= Dummy Then
                                    rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                                    QtyReadyStock1 -= Dummy
                                    Dummy = 0
                                Else
                                    Dummy -= QtyReadyStock1
                                    rowItemSO_SF.qty_ready_stock = QtyReadyStock1
                                    QtyReadyStock1 = 0

                                    'jika ada masih ada sisa dummy, maka kurangkan dengan qty po ready (dari pob baris sebelumnya)
                                    If QtyPOReady >= Dummy Then
                                        QtyPOReady -= Dummy
                                        Dummy = 0
                                    Else
                                        Dummy -= QtyPOReady
                                        QtyPOReady = 0
                                    End If
                                End If

                                'tambahkan qty used, pada pob baris ybs
                                rowItemPOB.qty_used += Dummy
                                If rowItemSO_SF.primary_po_ref = "" Then
                                    rowItemSO_SF.primary_po_ref = rowItemPOB.primary_po.ToString + ";" + Format(rowItemPOB.available_date, "dd MMM yyyy") _
                                    + ";" + Me.ObjCD.DigitGrouping(Dummy, rowItemSO_SF._Decimal)
                                Else
                                    rowItemSO_SF.primary_po_ref = rowItemSO_SF.primary_po_ref + "-" + rowItemPOB.primary_po.ToString + ";" + Format(rowItemPOB.available_date, "dd MMM yyyy") _
                                    + ";" + Me.ObjCD.DigitGrouping(Dummy, rowItemSO_SF._Decimal)
                                End If

                                'jika semua qty sudah habis dipakai, maka pindah ke baris item pob selanjutnya
                                If rowItemPOB.qty = rowItemPOB.qty_used Then
                                    IndexItemPOB += 1
                                End If
                            Else
                                'jika qty item delivery lebih besar dari qty po ready + qty po,ada 2 kemungkinan:
                                If IndexItemPOB + 1 < vTBLItemPOB.DefaultView.Count Then
                                    '1. tambahkan qty po dari baris selanjut, karena dari pob, maka qty ready stock tidak di tambahkan
                                    QtyPOReady += rowItemPOB.qty - rowItemPOB.qty_used
                                    'tandai primary po ref
                                    If rowItemSO_SF.primary_po_ref = "" Then
                                        rowItemSO_SF.primary_po_ref = rowItemPOB.primary_po.ToString + ";" + Format(rowItemPOB.available_date, "dd MMM yyyy") _
                                        + ";" + Me.ObjCD.DigitGrouping(rowItemPOB.qty - rowItemPOB.qty_used, rowItemSO_SF._Decimal)
                                    Else
                                        rowItemSO_SF.primary_po_ref = rowItemSO_SF.primary_po_ref + "-" + rowItemPOB.primary_po.ToString + ";" + Format(rowItemPOB.available_date, "dd MMM yyyy") _
                                        + ";" + Me.ObjCD.DigitGrouping(rowItemPOB.qty - rowItemPOB.qty_used, rowItemSO_SF._Decimal)
                                    End If

                                    rowItemPOB.qty_used = rowItemPOB.qty
                                    IndexItemPOB += 1
                                Else
                                    '2. tidak ada baris selanjutnya, maka qty availble hanya sejumlah yang ada,
                                    'estimate delivery tidak dikehui kapan, tapi available qty ada
                                    rowItemSO_SF.qty_available = rowItemPOB.qty - rowItemPOB.qty_used + QtyPOReady + QtyReadyStock1
                                    rowItemSO_SF.estimate_delivery = #1/1/2200#
                                    'perhitungan ready stock
                                    If QtyReadyStock1 >= rowItemSO_SF.qty_sob Then
                                        rowItemSO_SF.qty_ready_stock = rowItemSO_SF.qty_sob
                                        QtyReadyStock1 -= rowItemSO_SF.qty_sob
                                    Else
                                        rowItemSO_SF.qty_ready_stock = QtyReadyStock1
                                        QtyReadyStock1 = 0
                                    End If
                                    If rowItemSO_SF.primary_po_ref = "" Then
                                        rowItemSO_SF.primary_po_ref = rowItemPOB.primary_po.ToString + ";" + Format(rowItemPOB.available_date, "dd MMM yyyy") _
                                        + ";" + Me.ObjCD.DigitGrouping(rowItemPOB.qty - rowItemPOB.qty_used, rowItemSO_SF._Decimal)
                                    Else
                                        rowItemSO_SF.primary_po_ref = rowItemSO_SF.primary_po_ref + "-" + rowItemPOB.primary_po.ToString + ";" + Format(rowItemPOB.available_date, "dd MMM yyyy") _
                                        + ";" + Me.ObjCD.DigitGrouping(rowItemPOB.qty - rowItemPOB.qty_used, rowItemSO_SF._Decimal)
                                    End If

                                    'index item pob tetap dinaikan, untuk menghentikan looping perhitungan
                                    IndexItemPOB += 1
                                    'tandai bahwa semua item pob sudah dipergunakan
                                    rowItemPOB.qty_used = rowItemPOB.qty
                                    QtyPOReady = 0
                                End If
                            End If
                            'kembalikan posisi qty ready stock
                            If rowItemSO_SF.tipe <> CInt(ClsTransaksiShippingForecast_Item.EnumTipeItemSF.Purchase_Req) Then
                                QtyReadyStock = QtyReadyStock1
                            End If
                        End While
                    End If
                    'tambahkan qty po ready sisa ke dalam item sob terakhir
                    If QtyPOReady > 0 Then
                        If vTBLItemDelivery.DefaultView.Count - 1 >= IndexItemDelivery Then
                            rowItemSO_SF = vTBLItemDelivery.DefaultView(IndexItemDelivery).Row

                            rowItemSO_SF.qty_available = QtyPOReady
                            rowItemSO_SF.qty_ready_stock = QtyReadyStock
                            rowItemSO_SF.estimate_delivery = #1/1/2200#
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message + " - " + TBLPurchReqItemDelivery.DefaultView.RowFilter, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryCodeNo = Nothing
            DTA_ProductQty.Dispose()
            TBL_ProductQty.Dispose()
            DTAPOItem.Dispose()
            TBLPOItem.Dispose()
            TBL_POB.Dispose()
            ListSudah = Nothing
            DTAItemDeliveryPO.Dispose()
            TBLItemDeliveryPO.Dispose()
            DTAPackingList.Dispose()
            TBLPackingList.Dispose()
            DTASOItem.Dispose()
            TBLSOItem.Dispose()
            DTASOItemDelivery.Dispose()
            TBLSOItemDelivery.Dispose()
            DTAPendingDO.Dispose()
            TBLPendingDO.Dispose()
            DTABooking.Dispose()
            TBLBooking.Dispose()
            ObjPriorityLevel = Nothing
            FrmCek.Dispose()

            DTAPurchReqItem.Dispose()
            TBLPurchReqItem.Dispose()
            DTAPurchReqItemDelivery.Dispose()
            TBLPurchReqItemDelivery.Dispose()
        End Try
    End Sub
End Class
