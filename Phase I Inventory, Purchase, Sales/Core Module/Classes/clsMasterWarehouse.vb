Option Explicit On
Option Strict Off

Public Class clsMasterWarehouse
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mWarehouseCode As String = ""
    Dim mDescription As String = ""
    Dim mUseTransit As Integer = EnumUseTransitLocation.NotUseTransitLoc
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

    Public Enum EnumUseTransitLocation As Integer
        NotUseTransitLoc
        UseTransitLoc
    End Enum

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property WarehouseCode() As String
        Get
            Return Me.mWarehouseCode
        End Get
        Set(ByVal value As String)
            mWarehouseCode = value
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

    Public Property UseTransit() As Integer
        Get
            Return mUseTransit
        End Get
        Set(ByVal value As Integer)
            mUseTransit = value
        End Set
    End Property

    Public Property PrimaryTransitLocation() As Double
        Get
            Return mPrimaryTransitLocation
        End Get
        Set(ByVal value As Double)
            mPrimaryTransitLocation = value
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
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterWarehouse_SimpanBaru, vUserCode)

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
        Dim DTA As New dseMasterWarehouseTableAdapters.tblmaster_warehouseTableAdapter
        Dim ObjS As New clsMasterStorageArea

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.mWarehouseCode, Me.mDescription, vUserCode, Now, vUserCode, Now, Now, Me.mUseTransit, Me.mPrimaryTransitLocation) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mWarehouseCode).Value
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'update transit lokasi pada master lokasi
                If Me.mUseTransit = EnumUseTransitLocation.UseTransitLoc Then
                    ObjS.UpdateTransit_MasterLocation(Me.mPrimary, Me.mPrimaryTransitLocation)
                End If

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjS = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       EnumSearchCriteria) As Double
        Dim DTA As New dseMasterWarehouseTableAdapters.tblmaster_warehouse1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
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
            End If
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
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterWarehouse_SimpanPerubahan, vUserCode)

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
        Dim DTA As New dseMasterWarehouseTableAdapters.tblmaster_warehouseTableAdapter
        Dim ObjS As New clsMasterStorageArea

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Update(Me.mWarehouseCode, Me.mDescription, vUserCode, Now, Now, Me.mUseTransit, Me.mPrimaryTransitLocation, Me.mPrimary, Me.mTimestamp) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode

                'update transit lokasi pada master lokasi
                If Me.mUseTransit = EnumUseTransitLocation.UseTransitLoc Then
                    ObjS.UpdateTransit_MasterLocation(Me.mPrimary, Me.mPrimaryTransitLocation)
                End If

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjS = Nothing
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
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterWarehouse_Hapus, vUser)

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
        Dim DTA As New dseMasterWarehouseTableAdapters.tblmaster_warehouseTableAdapter

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

    Public Function ApakahSudahAda(ByVal vWarehouseCode As String) As Boolean
        Dim DTA As New dseMasterWarehouseTableAdapters.tblmaster_warehouseTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vWarehouseCode).HasValue Then
                If DTA.ScalarQuery_GetJml(vWarehouseCode).Value > 0 Then
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

    Public Function ApakahBolehHapus(ByVal vPrimary As Double) As Boolean
        Dim Obj As New clsMasterStorageArea
        Dim Lvl_Int As Integer = 0
        Dim DTA1 As New dseMasterStorageAreaTableAdapters.tbltrans_material_planning_item_location_exclusionTableAdapter
        Dim TBL1 As New dseMasterStorageArea.tbltrans_material_planning_item_location_exclusionDataTable 
        Dim ObjCustom As New ClsIdentifier

        Try
            'tidak boleh dihapus bila sudah ada yang merefer:
            '1. storage area
            If Obj.GetJml_ByWarehouse(vPrimary) > 0 Then
                MsgBox("You cannot delete this data, because there are storage area using this data.", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            'jika sudah digunakan pada material exclusion
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Fill(TBL1, vPrimary, Lvl_Int)
            If TBL1.Rows.Count > 0 Then
                MsgBox("You cannot delete this data, there are material planning transaction refer to this data.", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            If ObjCustom.GetIdentifier = "smess_lmb" Then
                If Me.ApakahBolehHapus_CustomLMB(vPrimary) = False Then
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
            DTA1.Dispose()
            TBL1.Dispose()
            ObjCustom = Nothing
        End Try
    End Function

    Public Function ApakahBolehHapus_CustomLMB(ByVal vPrimaryWarehouse As Double) As Boolean
        Dim DTA As New dseMasterWarehouseTableAdapters.DataTable1TableAdapter
        Dim TBL As New dseMasterWarehouse.DataTable1DataTable
        Dim row As dseMasterWarehouse.DataTable1Row
        Dim Msg As String = ""
        Dim counter As Integer = 0

        Try
            'tidak boleh hapus jika sudah di rujuk pada master setting bundling
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryWarehouse)
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    counter += 1
                    If Msg = "" Then
                        Msg = counter.ToString + ". " + row.Code
                    Else
                        Msg = Msg + vbCrLf + counter.ToString + ". " + row.Code
                    End If
                Next

                MsgBox("This data is referenced by bundling warehouse setting for the following process:" + vbCrLf + Msg + _
                vbCrLf + "You can't delete this data.", MsgBoxStyle.Information, "Information")

                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As dseMasterWarehouse.tblmaster_warehouse1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New dseMasterWarehouseTableAdapters.tblmaster_warehouse1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Warehouse code"
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
        Dim DTA As New dseMasterWarehouseTableAdapters.tblmaster_warehouseTableAdapter
        Dim TBL As New dseMasterWarehouse.tblmaster_warehouseDataTable
        Dim row As dseMasterWarehouse.tblmaster_warehouseRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy(TBL, Primary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = row.Primary
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mDescription = .Description
                    If .Use_Transit = 0 Then
                        Me.mUseTransit = EnumUseTransitLocation.NotUseTransitLoc
                    Else
                        Me.mUseTransit = EnumUseTransitLocation.UseTransitLoc
                    End If
                    Me.mPrimaryTransitLocation = .Primary_Transit_Location
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mWarehouseCode = .Warehouse_Code
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

    Public Function GetWarehouseCode(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterWarehouseTableAdapters.tblmaster_warehouseTableAdapter

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

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New dseMasterWarehouseTableAdapters.tblmaster_warehouseTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetTimestamp(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vWarehouseCode As String) As Double
        Dim DTA As New dseMasterWarehouseTableAdapters.tblmaster_warehouseTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vWarehouseCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vWarehouseCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDescription(ByVal vPrimary As Double) As String
        Dim DTA As New dseMasterWarehouseTableAdapters.tblmaster_warehouseTableAdapter
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetDesc(vPrimary)
            If String.IsNullOrEmpty(Retval) = False Then
                Return Retval
            Else
                Return Retval
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetIsTransit(ByVal vPrimaryWarehouse As Double, ByRef vPrimaryTransitLoc As Double) As Boolean
        Dim dta As New dseMasterWarehouseTableAdapters.tblmaster_warehouse2TableAdapter
        Dim tbl As New dseMasterWarehouse.tblmaster_warehouse2DataTable
        Dim row As dseMasterWarehouse.tblmaster_warehouse2Row
        Dim UseTransit_ As Integer = 0
        Dim PrimaryTransitLoc_ As Double = -1

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(tbl, vPrimaryWarehouse)
            If tbl.Count > 0 Then
                row = tbl.Rows(0)
                With row
                    UseTransit_ = .Use_Transit
                    PrimaryTransitLoc_ = .Primary_Transit_Location
                    If UseTransit_ <> 0 Then
                        vPrimaryTransitLoc = PrimaryTransitLoc_
                        Return True
                    End If
                End With
            End If

            If UseTransit_ = 0 Then
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
End Class
