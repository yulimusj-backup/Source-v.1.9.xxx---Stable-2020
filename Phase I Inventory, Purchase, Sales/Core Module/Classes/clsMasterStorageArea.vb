Option Explicit On
Option Strict Off
Public Class clsMasterStorageArea
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mStorageAreaCode As String = ""
    Dim mDescription As String = ""
    Dim mPrimaryParent As Double = -1
    Dim mIsTransit As Integer = enumTransit.NotTransit
    Dim mPrimaryTransitLocation As Double = -1
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    'digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Public Enum enumTransit As Integer
        NotTransit
        Transit
        All
    End Enum

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property StorageAreaCode() As String
        Get
            Return Me.mStorageAreaCode
        End Get
        Set(ByVal value As String)
            mStorageAreaCode = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return Me.mDescription
        End Get
        Set(ByVal value As String)
            mDescription = value
        End Set
    End Property

    Public Property IsTransit() As Double
        Get
            Return Me.mIsTransit
        End Get
        Set(ByVal value As Double)
            mIsTransit = value
        End Set
    End Property

    Public Property PrimaryTransitLocation() As Double
        Get
            Return Me.mPrimaryTransitLocation
        End Get
        Set(ByVal value As Double)
            mPrimaryTransitLocation = value
        End Set
    End Property

    Public Property PrimaryParent() As Double
        Get
            Return Me.mPrimaryParent
        End Get
        Set(ByVal value As Double)
            mPrimaryParent = value
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

    Public Overridable Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterStorageArea_SimpanBaru, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vUserCode)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Function SimpanBaru1(ByVal vUserCode As String) As Boolean
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            If Me.ApakahDatabenar() = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                If DTA.Insert(Me.mPrimaryParent, Me.mStorageAreaCode, Me.mDescription, _
                vUserCode, Now, vUserCode, Now, Now, Me.mIsTransit, Me.mPrimaryTransitLocation) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mStorageAreaCode).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    Return True
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterStorageArea_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vUserCode)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Function SimpanPerubahan1(ByVal vUserCode As String) As Boolean
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            If Me.ApakahDatabenar() = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                If DTA.Update(Me.mPrimaryParent, Me.mStorageAreaCode, Me.mDescription, _
                vUserCode, Now, Now, Me.mIsTransit, Me.mPrimaryTransitLocation, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    Return True
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUser As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterStorageArea_Hapus, vUser)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Function Hapus1(ByVal vPrimary As Double) As Boolean
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary) = True Then
                DTA.Delete(vPrimary)
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

    Public Function ApakahSudahAda(ByVal vStorageAreaCode As String, Optional ByVal vEnumTransit As enumTransit = enumTransit.NotTransit) As Boolean
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter
        Dim Value_ As Nullable(Of Long)
        Dim Transit_0, Transit_1 As Integer

        Try
            If vEnumTransit = enumTransit.NotTransit Then
                Transit_0 = 0
                Transit_1 = 0
            ElseIf vEnumTransit = enumTransit.Transit Then
                Transit_0 = 1
                Transit_1 = 1
            Else
                Transit_0 = 0
                Transit_1 = 1
            End If

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Value_ = DTA.ScalarQuery_GetJml(vStorageAreaCode, Transit_0, Transit_1)
            If Value_.HasValue Then
                If Value_.Value > 0 Then
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

    Public Function ApakahBolehUbahIsTransit(ByVal vPrimary As Double) As Boolean
        Try
            If Me.ApakahBolehHapus(vPrimary, False) = False Then
                MsgBox("You cannot change is transit, there are other transaction refer to this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahDatabenar() As Boolean
        Dim ObjW As New clsMasterWarehouse
        Dim PrimaryDummy As Double = Me.mPrimaryTransitLocation

        Try
            If Me.mIsTransit = 1 Then
                Me.mPrimaryTransitLocation = -1
            Else
                ObjW.GetIsTransit(Me.mPrimaryParent, Me.mPrimaryTransitLocation)
                If PrimaryDummy <> -1 And Me.mPrimaryTransitLocation = -1 Then
                    MsgBox("Transit location data changed to maintain data consistency with warehouse master data.", MsgBoxStyle.Information, "Information")
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjW = Nothing
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, Optional ByVal ShowMsgBox As Boolean = True) As Boolean
        Dim Obj As New clsMasterProduct

        'refer to: FrmTrans_MaterialPlanningDashboard_Adv_Exclusion
        Dim Lvl_Int As Integer = 1
        Dim DTA1 As New dseMasterStorageAreaTableAdapters.tbltrans_material_planning_item_location_exclusionTableAdapter
        Dim TBL1 As New dseMasterStorageArea.tbltrans_material_planning_item_location_exclusionDataTable
        Dim DTA2 As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            'tidak boleh dihapus bila sudah ada yang merefer:
            '1. product 
            If Obj.GetJml_ByStorageArea(vPrimary) > 0 Then
                If ShowMsgBox = True Then
                    MsgBox("You cannot delete this data, there are product(s) stored in this area.", MsgBoxStyle.Information, "Information")
                End If

                Return False
            End If
            Obj = Nothing

            'jika sudah digunakan pada stock in out
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA2.ScalarQuery_GetJmlStockInOut(vPrimary)
            If RetVal.HasValue = True Then
                If RetVal.Value > 0 Then
                    If ShowMsgBox = True Then
                        MsgBox("You cannot delete this data, there are stock in/out transaction refer to this data.", MsgBoxStyle.Information, "Information")
                    End If


                    Return False
                End If
            End If

            'jika sudah digunakan pada location transfer
            RetVal = DTA2.ScalarQuery_GetJml_StockTransfer(vPrimary)
            If RetVal.HasValue = True Then
                If RetVal.Value > 0 Then
                    If ShowMsgBox = True Then
                        MsgBox("You cannot delete this data, there are location transfer refer to this data.", MsgBoxStyle.Information, "Information")
                    End If

                    Return False
                End If
            End If

            'jika sudah digunakan pada item before bundling
            RetVal = DTA2.ScalarQuery_GetJml_StockBundling_Before(vPrimary)
            If RetVal.HasValue = True Then
                If RetVal.Value > 0 Then
                    If ShowMsgBox = True Then
                        MsgBox("You cannot delete this data, there are bundling transaction refer to this data.", MsgBoxStyle.Information, "Information")
                    End If

                    Return False
                End If
            End If

            'jika sudah digunakan pada item after bundling
            RetVal = DTA2.ScalarQuery_GetJml_StockBundling_After(vPrimary)
            If RetVal.HasValue = True Then
                If RetVal.Value > 0 Then
                    If ShowMsgBox = True Then
                        MsgBox("You cannot delete this data, there are bundling transaction refer to this data.", MsgBoxStyle.Information, "Information")
                    End If

                    Return False
                End If
            End If

            'jika sudah digunakan pada material exclusion
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Fill(TBL1, vPrimary, Lvl_Int)
            If TBL1.Rows.Count > 0 Then
                If ShowMsgBox = True Then
                    MsgBox("You cannot delete this data, there are material planning transaction refer to this data.", MsgBoxStyle.Information, "Information")
                End If

                Return False
            End If

            'Untuk Storage Area Yang IsTransit
            'jika ubah dari is transit ke not is transit atau hapus maka cek apa storage area nya sudah di rujuk oleh :
            '1 storage area yang bukan transit
            If Me.GetJml_TransitInLocation(vPrimary) > 0 Then
                If ShowMsgBox = True Then
                    MsgBox("You cannot delete this data, there are other storage area refer to this data.", MsgBoxStyle.Information, "Information")
                End If

                Return False
            End If
            '2 warehouse
            If Me.GetJml_TransitInWarehouse(vPrimary) > 0 Then
                If ShowMsgBox = True Then
                    MsgBox("You cannot delete this data, there are warehouse refer to this data.", MsgBoxStyle.Information, "Information")
                End If

                Return False
            End If
            '3 stock location transfer item
            If Me.GetJml_TransitInStockTransfer(vPrimary) > 0 Then
                If ShowMsgBox = True Then
                    MsgBox("You cannot delete this data, there are stock transfer transaction refer to this data.", MsgBoxStyle.Information, "Information")
                End If

                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
            DTA1.Dispose()
            TBL1.Dispose()
            DTA2.Dispose()
        End Try
    End Function

    Public Sub IsiTBL(ByVal vTBL As dseMasterStorageArea.tblmaster_storage_area1DataTable)
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_area1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vTBL.Rows.Clear()

            DTA.Fill(vTBL)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As dseMasterStorageArea.DataTable1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal vEnumTransit As enumTransit) As Boolean
        Dim DTA As New dseMasterStorageAreaTableAdapters.DataTable1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Dim Transit_0, Transit_1 As Integer

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If vEnumTransit = enumTransit.NotTransit Then
                Transit_0 = 0
                Transit_1 = 0
            ElseIf vEnumTransit = enumTransit.Transit Then
                Transit_0 = 1
                Transit_1 = 1
            Else
                Transit_0 = 0
                Transit_1 = 1
            End If

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%", Transit_0, Transit_1)
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%", Transit_0, Transit_1)
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

    Public Function CariDetail(ByVal Primary As Double) As Boolean
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter
        Dim TBL As New dseMasterStorageArea.tblmaster_storage_areaDataTable
        Dim row As dseMasterStorageArea.tblmaster_storage_areaRow

        Try
            DTA.FillBy_Primary(TBL, Primary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = row.Primary
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mDescription = .Description
                    If .Is_Transit = 0 Then
                        Me.mIsTransit = enumTransit.NotTransit
                    Else
                        Me.mIsTransit = enumTransit.Transit
                    End If
                    Me.mPrimaryTransitLocation = .Primary_Transit_Location
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mPrimaryParent = .Primary_Warehouse
                    Me.mStorageAreaCode = .Storage_Area_Code
                    Me.mTimestamp = .Timestamp
                End With
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan jumlah storage area yang ada
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJml() As Double
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlTotal.HasValue = True Then
                Return DTA.ScalarQuery_GetJmlTotal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_ByCode(ByVal vCode As String, Optional ByVal vEnumTransit As enumTransit = enumTransit.NotTransit) As Double
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter
        Dim RetVal As Nullable(Of Long)
        Dim Transit_0, Transit_1 As Integer

        Try
            If vEnumTransit = enumTransit.NotTransit Then
                Transit_0 = 0
                Transit_1 = 0
            ElseIf vEnumTransit = enumTransit.Transit Then
                Transit_0 = 1
                Transit_1 = 1
            Else
                Transit_0 = 0
                Transit_1 = 1
            End If

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJml(vCode, Transit_0, Transit_1)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function GetJmlBy_LocationCodeLIKE(ByVal vLocationCodeLIKE As String, Optional ByVal vEnumTransit As enumTransit = enumTransit.NotTransit) As Double
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter
        Dim RetVal As Nullable(Of Long)
        Dim Transit_0, Transit_1 As Integer

        Try
            If vEnumTransit = enumTransit.NotTransit Then
                Transit_0 = 0
                Transit_1 = 0
            ElseIf vEnumTransit = enumTransit.Transit Then
                Transit_0 = 1
                Transit_1 = 1
            Else
                Transit_0 = 0
                Transit_1 = 1
            End If

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJml_ByCodeLIKE(vLocationCodeLIKE, Transit_0, Transit_1)
            If RetVal.HasValue = True Then
                Return RetVal.Value
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
    EnumSearchCriteria, ByVal vEnumTransit As enumTransit) As Double
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter
        Dim Transit_0, Transit_1 As Integer

        Try
            If vEnumTransit = enumTransit.NotTransit Then
                Transit_0 = 0
                Transit_1 = 0
            ElseIf vEnumTransit = enumTransit.Transit Then
                Transit_0 = 1
                Transit_1 = 1
            Else
                Transit_0 = 0
                Transit_1 = 1
            End If

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                If DTA.ScalarQuery_GetJml1(vKriteria + "%", Transit_0, Transit_1).HasValue Then
                    Return DTA.ScalarQuery_GetJml1(vKriteria + "%", Transit_0, Transit_1).Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                If DTA.ScalarQuery_GetJml2(vKriteria + "%", Transit_0, Transit_1).HasValue Then
                    Return DTA.ScalarQuery_GetJml2(vKriteria + "%", Transit_0, Transit_1).Value
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

    Public Function GetDescription(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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

    Public Function GetStorageAreaCode(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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

    Public Function GetJml_ByWarehouse(ByVal vPrimaryWarehouse As Double) As Integer
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter
        Dim Jml As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Jml = DTA.ScalarQuery_GetJml3(vPrimaryWarehouse)
            If Jml.HasValue Then
                Return Jml.Value
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
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetTimestamp(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetWarehouseDesc(ByVal vPrimaryStorageArea As Double) As String
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetWarehouseDesc(vPrimaryStorageArea)) = False Then
                Return DTA.ScalarQuery_GetWarehouseDesc(vPrimaryStorageArea)
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

    Public Function GetWarehouseCode(ByVal vPrimaryStorageArea As Double) As String
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetWarehouseCode(vPrimaryStorageArea)) = False Then
                Return DTA.ScalarQuery_GetWarehouseCode(vPrimaryStorageArea)
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

    Public Function GetPrimary(ByVal vStorageAreaCode As String) As Double
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vStorageAreaCode).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary(vStorageAreaCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryWarehouse(ByVal vPrimary As Double) As Double
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimaryWarehouse(vPrimary)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function UpdateTransit_MasterLocation(ByVal vPrimaryWarehouse As Double,  ByVal vPrimaryTransitLoc As Double) As Boolean
        Dim dta As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.UpdateQuery_Transit(Now, vPrimaryTransitLoc, vPrimaryWarehouse, enumTransit.NotTransit) > 0 Then
                Return True
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function GetJml_TransitInLocation(ByVal vPrimaryTransitLocation As Double) As Integer
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter
        Dim retval As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            retval = DTA.ScalarQuery_GetJmlTransitLoc(vPrimaryTransitLocation)
            If retval.HasValue Then
                Return retval.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_TransitInWarehouse(ByVal vPrimaryTransitLocation As Double) As Integer
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter
        Dim retval As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            retval = DTA.ScalarQuery_GetJmlTransitLoc_InWarehouse(vPrimaryTransitLocation)
            If retval.HasValue Then
                Return retval.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_TransitInStockTransfer(ByVal vPrimaryTransitLocation As Double) As Integer
        Dim DTA As New dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter
        Dim retval As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            retval = DTA.ScalarQuery_GetJmlTransitLoc_InStockTransfer(vPrimaryTransitLocation)
            If retval.HasValue Then
                Return retval.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTransitData(ByVal vPrimaryStorageArea As Double, ByRef vPrimaryTransitLoc As Double, _
    ByRef vTransitLoc As String, ByRef vTransitWarehouse As String) As Boolean
        Dim dta As New dseMasterStorageAreaTableAdapters.tbl_transitTableAdapter
        Dim tbl As New dseMasterStorageArea.tbl_transitDataTable
        Dim row As dseMasterStorageArea.tbl_transitRow

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.FillBy_PrimaryStorageArea(tbl, vPrimaryStorageArea)
            If tbl.Count > 0 Then
                row = tbl.Rows(0)
                With row
                    If .Is_Transit = 0 Then
                        'jika menggunakan location transit
                        vPrimaryTransitLoc = .Primary_Transit_Location
                        vTransitLoc = .Transit_Code
                        vTransitWarehouse = .Transit_Warhouse_Code
                        Return True
                    End If
                End With
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
End Class
