Option Explicit On
Option Strict Off

Public Class ClsUser_TempPass
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mPrimaryUser As Double = -1
    Dim mRemarks As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianMaster
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property PrimaryUser() As Double
        Get
            Return mPrimaryUser
        End Get
        Set(ByVal value As Double)
            mPrimaryUser = value
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

    Public Function ApakahSudahAda(ByVal vPrimaryUser As Double) As Boolean
        Dim DTA As New DSUser_TempPassTableAdapters.tblmaster_user_temppassTableAdapter
        Dim Retval As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJmlByPrimaryUser(vPrimaryUser)
            If Retval.HasValue = True Then
                If Retval.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vDgvItem As DataGridView, ByVal vFrm As FrmMasterUser_TempPass) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterUser_TempPass_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vUserCode, vDgvItem, vFrm)
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

    Private Function SimpanBaru1(ByVal vUserCode As String, ByVal vDgvItem As DataGridView, ByVal vFrm As FrmMasterUser_TempPass) As Boolean
        Dim DTA As New DSUser_TempPassTableAdapters.tblmaster_user_temppassTableAdapter
        Dim DTAItem As New DSUser_TempPassTableAdapters.tblmaster_user_temppass_itemTableAdapter
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DummyBool As Boolean = False

        Try
            If Me.ApakahSudahAda(Me.mPrimaryUser) = True Then
                MsgBox("User code duplication, please change user code.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = DTA.Connection.ConnectionString

            If DTA.InsertQuery(Me.mPrimaryUser, _
            Me.mRemarks, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimaryByPrimaryUser(Me.mPrimaryUser).Value
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'simpan item
                DummyBool = vDgvItem.AllowUserToAddRows
                vDgvItem.AllowUserToAddRows = False
                For Each rowDGV As DataGridViewRow In vDgvItem.Rows
                    DTAItem.InsertQuery(Me.mPrimary, CDate(rowDGV.Cells(vFrm.ColStartDate.Index).Value), CDate(rowDGV.Cells(vFrm.ColEndDate.Index).Value), _
                    ObjCD.Convert_ToMinutes(rowDGV.Cells(vFrm.ColStartTime.Index).Value), ObjCD.Convert_ToMinutes(rowDGV.Cells(vFrm.ColEndTime.Index).Value), _
                    rowDGV.Cells(vFrm.ColRemarks.Index).Value, rowDGV.Cells(vFrm.ColPass.Index).Value)
                Next
                vDgvItem.AllowUserToAddRows = DummyBool

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal vDgvItem As DataGridView, ByVal vFrm As FrmMasterUser_TempPass) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterUser_TempPass_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vUserCode, vDgvItem, vFrm)
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
            ObjAntri = Nothing
            SW = Nothing
        End Try
    End Function

    Private Function SimpanPerubahan1(ByVal vUserCode As String, ByVal vDgvItem As DataGridView, ByVal vFrm As FrmMasterUser_TempPass) As Boolean
        Dim DTA As New DSUser_TempPassTableAdapters.tblmaster_user_temppassTableAdapter
        Dim DTAItem As New DSUser_TempPassTableAdapters.tblmaster_user_temppass_itemTableAdapter
        Dim RetVal As Nullable(Of Long)
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DummyBool As Boolean = False

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = DTA.Connection.ConnectionString

            'cek apakah ada transaksi lain yang memiliki primary user yang sama
            RetVal = DTA.ScalarQuery_GetJml2(Me.mPrimary, Me.mPrimaryUser)
            If RetVal.HasValue = True Then
                If RetVal.Value > 0 Then
                    MsgBox("User code duplication, please change user code.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If DTA.UpdateQuery(Me.mPrimaryUser, Me.mRemarks, _
            vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode

                'hapuskan data item lama
                DTAItem.DeleteQuery(Me.mPrimary)

                'simpan data item baru
                DummyBool = vDgvItem.AllowUserToAddRows
                vDgvItem.AllowUserToAddRows = False
                For Each rowDGV As DataGridViewRow In vDgvItem.Rows
                    DTAItem.InsertQuery(Me.mPrimary, CDate(rowDGV.Cells(vFrm.ColStartDate.Index).Value), CDate(rowDGV.Cells(vFrm.ColEndDate.Index).Value), _
                    ObjCD.Convert_ToMinutes(rowDGV.Cells(vFrm.ColStartTime.Index).Value), ObjCD.Convert_ToMinutes(rowDGV.Cells(vFrm.ColEndTime.Index).Value), _
                    rowDGV.Cells(vFrm.ColRemarks.Index).Value, rowDGV.Cells(vFrm.ColPass.Index).Value)
                Next
                vDgvItem.AllowUserToAddRows = DummyBool

                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal vTampilkanPesan As Boolean, ByVal vStrictDelete As Boolean) As Boolean
        Dim DTA As New DSUser_TempPassTableAdapters.tblmaster_user_temppassTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vStrictDelete, vUserCode, vTampilkanPesan) = True Then
                DTA.DeleteQuery(vPrimary)
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

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vStrictDelete As Boolean, ByVal vUserCode As String, ByVal vTampilkanPesan As Boolean) As Boolean
        Dim DTA As New DSUser_TempPassTableAdapters.tblmaster_user_temppassTableAdapter
        Dim vCreatedBy As String = ""
        Dim ObjUser As New Core_Module.ClsUser

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            '1. jika strict delete
            If vStrictDelete = True Then
                vCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                If String.IsNullOrEmpty(vCreatedBy) = False Then
                    If vCreatedBy <> vUserCode Then
                        If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmMasterUser_TempPass_OverrideDelete") = False Then
                            If vTampilkanPesan = True Then
                                MsgBox("Only user: " + vCreatedBy + " allowed to delete this data.", MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjUser = Nothing
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vStrictDelete As Boolean, ByVal vUserCode As String, ByVal vTampilkanPesan As Boolean) As Boolean
        Dim DTA As New DSUser_TempPassTableAdapters.tblmaster_user_temppassTableAdapter
        Dim vCreatedBy As String = ""
        Dim ObjUser As New Core_Module.ClsUser

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            '1. jika strict delete
            If vStrictDelete = True Then
                vCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                If String.IsNullOrEmpty(vCreatedBy) = False Then
                    If vCreatedBy <> vUserCode Then
                        If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmMasterUser_TempPass_OverrideEdit") = False Then
                            If vTampilkanPesan = True Then
                                MsgBox("Only user: " + vCreatedBy + " allowed to edit this data.", MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjUser = Nothing
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DSUser_TempPass.tblmaster_user_temppass1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DSUser_TempPassTableAdapters.tblmaster_user_temppass1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_UserCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_UserNameLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
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

    Public Function CariDetail(ByVal Primary As Double, ByVal vDgvItem As DataGridView, ByVal vFrm As FrmMasterUser_TempPass) As Boolean
        Dim DTA As New DSUser_TempPassTableAdapters.tblmaster_user_temppassTableAdapter
        Dim TBL As New DSUser_TempPass.tblmaster_user_temppassDataTable
        Dim row As DSUser_TempPass.tblmaster_user_temppassRow

        Dim DTAItem As New DSUser_TempPassTableAdapters.tblmaster_user_temppass_itemTableAdapter
        Dim TBLItem As New DSUser_TempPass.tblmaster_user_temppass_itemDataTable
        Dim rowItem As DSUser_TempPass.tblmaster_user_temppass_itemRow
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DummyBool As Boolean = False

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.FillBy_Primary(TBL, Primary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = row.primary
                    Me.mPrimaryUser = row.primary_user
                    Me.mRemarks = .remarks

                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With
            End If

            'isikan data item
            vDgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgvItem.Rows.Clear()
            DTAItem.FillBy_PrimaryTransaksi(TBLItem, Me.mPrimary)
            DummyBool = vDgvItem.AllowUserToAddRows
            vDgvItem.AllowUserToAddRows = False
            For Each rowItem In TBLItem.Rows
                vDgvItem.Rows.Add()
                rowDGV = vDgvItem.Rows(vDgvItem.Rows.Count - 1)
                With rowDGV
                    .Cells(vFrm.ColConfirmPassword.Index).Value = rowItem.password
                    .Cells(vFrm.ColEndDate.Index).Value = Format(rowItem.end_date, "dd MMM yyyy")
                    .Cells(vFrm.ColEndTime.Index).Value = ObjCD.Convert_ToString(rowItem.end_time)
                    .Cells(vFrm.ColNo.Index).Value = vDgvItem.RowCount
                    .Cells(vFrm.ColPass.Index).Value = rowItem.password
                    .Cells(vFrm.ColRemarks.Index).Value = rowItem.remarks
                    .Cells(vFrm.ColStartDate.Index).Value = Format(rowItem.start_date, "dd MMM yyyy")
                    .Cells(vFrm.ColStartTime.Index).Value = ObjCD.Convert_ToString(rowItem.start_time)
                    .Cells(vFrm.ColConfirmPassword_Show.Index).Value = "*****"
                    .Cells(vFrm.ColPass_Show.Index).Value = "*****"
                End With
            Next
            vDgvItem.AllowUserToAddRows = DummyBool
            vDgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Function

    Public Sub ResetObject(ByVal vDGVItem As DataGridView)
        Try
            Me.mCreatedBy = ""
            Me.mCreatedDate = Now
            Me.mModifiedBy = ""
            Me.mModifiedDate = Now
            Me.mPrimary = -1
            Me.mPrimaryUser = -1
            Me.mRemarks = ""
            Me.mTimestamp = Now
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetJmlTranskasiLainByPrimaryUser(ByVal vPrimaryUser As Double, ByVal vPrimaryTransaksi As Double) As Double
        Dim DTA As New DSUser_TempPassTableAdapters.tblmaster_user_temppassTableAdapter
        Dim Retval As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJml2(vPrimaryTransaksi, vPrimaryUser)
            If Retval.HasValue = True Then
                Return Retval.Value
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
    EnumSearchCriteria) As Double
        Dim DTA As New DSUser_TempPassTableAdapters.tblmaster_user_temppass1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                If DTA.ScalarQuery_GetJmlByUserCodeLIKE(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByUserCodeLIKE(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                If DTA.ScalarQuery_GetJmlByUserNameLIKE(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByUserNameLIKE(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                If DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%").Value
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

    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DSUser_TempPassTableAdapters.tblmaster_user_temppassTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetTimestamp(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Login(ByVal vKodeUser As String, ByVal vPass As String) As Boolean
        Dim ObjUser As New ClsUser
        Dim PrimaryUser1 As Double = -1
        Dim DTAItem As New DSUser_TempPassTableAdapters.tblmaster_user_temppass_itemTableAdapter
        Dim TBLItem As New DSUser_TempPass.tblmaster_user_temppass_itemDataTable
        Dim rowItem As DSUser_TempPass.tblmaster_user_temppass_itemRow
        Dim LoginTime As Date = Now
        Dim LoginTime_Min As Double = 0
        Dim ObjCD As New Core_Module.ClsCekData
        Dim MM As String = ""
        Dim HH As String = ""

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            PrimaryUser1 = ObjUser.GetPrimary(vKodeUser)

            DTAItem.FillBy_PrimaryUser(TBLItem, PrimaryUser1, vPass)
            For Each rowItem In TBLItem.Rows
                If rowItem.start_time <= rowItem.end_time Then
                    If rowItem.start_date.Date <= LoginTime.Date And LoginTime.Date <= rowItem.end_date.Date Then
                        LoginTime_Min = ObjCD.Convert_ToMinutes(Format(LoginTime, "HH.mm"))
                        If LoginTime_Min >= rowItem.start_time And LoginTime_Min <= rowItem.end_time Then
                            Return True
                        End If
                    End If
                Else
                    LoginTime_Min = ObjCD.Convert_ToMinutes(Format(LoginTime, "HH.mm"))
                    If LoginTime_Min >= rowItem.start_time Then
                        If rowItem.start_date.Date <= LoginTime.Date And LoginTime.Date <= rowItem.end_date.Date Then
                            Return True
                        End If
                    Else
                        If LoginTime_Min <= rowItem.end_time Then
                            If (rowItem.start_date.Date + TimeSpan.FromDays(1)).Date <= LoginTime.Date And _
                            LoginTime.Date <= (rowItem.end_date.Date + TimeSpan.FromDays(1)).Date Then
                                Return True
                            End If
                        End If
                    End If
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjUser = Nothing
            DTAItem.Dispose()
            TBLItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Sub New()
        Me.DTAInOUT.Connection.ConnectionString = Core_Module.GetConstring
    End Sub
End Class
