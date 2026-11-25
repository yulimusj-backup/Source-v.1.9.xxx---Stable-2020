Option Explicit On
Option Strict Off

Public Class ClsMasterKompetensi
    'default penulisan value dalam bentk string:
    '1. number      : menggunakan fungsi obj digit group dan menggunakan default prec
    '2. tanggal     : format MM/dd/yyyy
    '3. text        : sesuai dengan textnya
    '4. YN          : false -> 0, true -> 1
    '5. Alfabetic   : A -> 0, B -> 1, C -> 2, dst

    Private mPrimary As Double = -1
    Private mDataCode As String = ""
    Private mDesc As String = ""
    Private mDataType As EnumDataType = EnumDataType.Number
    Private mDefaultPrecision As Integer = 0
    Private mValue As String = ""
    Private mRemarks As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = ClsAntrian.EnumTipeAntrian.AntrianMaster
    Dim ObjRnd As New Random
    Dim ObjAntri As New Core_Module.ClsAntrian
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New DseAntrianTableAdapters.tbltrans_antrianTableAdapter

    Public Enum EnumDataType As Integer
        Number
        Tanggal
        Text
        YN
        Alfabetic
    End Enum

    Public Enum Aphabetic As Integer
        A
        B
        C
        D
        E
        F
    End Enum

    Public Property DataType() As EnumDataType
        Get
            Return mDataType
        End Get
        Set(ByVal value As EnumDataType)
            mDataType = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return mDesc
        End Get
        Set(ByVal value As String)
            mDesc = value
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

    Public Property DataCode() As String
        Get
            Return mDataCode
        End Get
        Set(ByVal value As String)
            mDataCode = value
        End Set
    End Property

    Public Property DefaultPrecision() As Integer
        Get
            Return mDefaultPrecision
        End Get
        Set(ByVal value As Integer)
            mDefaultPrecision = value
        End Set
    End Property

    Public Property Value() As String
        Get
            Return mValue
        End Get
        Set(ByVal value As String)
            mValue = value
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

    Public Sub New()
        Try
            Me.DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetDefPrec(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetDefaultPrecision(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetDefaultPrecision(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDataType(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetDataType(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetDataType(vPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDesc(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetDesc(vPrimary)

            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
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

    Public Function GetValue(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetValue(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetValue(vPrimary)
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

    Public Function SetFormatDataCustom(ByVal vPar As String, ByVal vDataType As ClsMasterKompetensi.EnumDataType, _
    ByVal vDef_Prec As Integer, ByVal vObjCD As Core_Module.ClsCekData) As String
        Try
            If vDataType = ClsMasterKompetensi.EnumDataType.Number Then
                Return vObjCD.DigitGrouping(vPar, vDef_Prec)
            ElseIf vDataType = ClsMasterKompetensi.EnumDataType.Tanggal Then
                Return Format(CDate(vPar), "dd MMM yyyy")
            ElseIf vDataType = ClsMasterKompetensi.EnumDataType.Text Then
                Return vPar
            ElseIf vDataType = ClsMasterKompetensi.EnumDataType.YN Then
                If vPar = "1" Then
                    Return "Yes"
                Else
                    Return "No"
                End If
            ElseIf vDataType = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                If vPar = ClsMasterKompetensi.Aphabetic.A Then
                    Return "A"
                ElseIf vPar = ClsMasterKompetensi.Aphabetic.B Then
                    Return "B"
                ElseIf vPar = ClsMasterKompetensi.Aphabetic.C Then
                    Return "C"
                ElseIf vPar = ClsMasterKompetensi.Aphabetic.D Then
                    Return "D"
                ElseIf vPar = ClsMasterKompetensi.Aphabetic.E Then
                    Return "E"
                ElseIf vPar = ClsMasterKompetensi.Aphabetic.F Then
                    Return "F"
                End If
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function SetFormatDataCustom(ByVal vPar As String, ByVal vPrimaryKompetensi As Double, ByVal vObjCD As Core_Module.ClsCekData) As String
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Dim vDataType As ClsMasterKompetensi.EnumDataType
        Dim vDef_Prec As Integer
        Dim TBL As New DseMasterKompetensi.tblmaster_kompetensiDataTable
        Dim row As DseMasterKompetensi.tblmaster_kompetensiRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, vPrimaryKompetensi)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                vDataType = row.Data_Type
                vDef_Prec = row.Default_Precision

                If vDataType = ClsMasterKompetensi.EnumDataType.Number Then
                    Return vObjCD.DigitGrouping(vPar, vDef_Prec)
                ElseIf vDataType = ClsMasterKompetensi.EnumDataType.Tanggal Then
                    Return Format(CDate(vPar), "dd MMM yyyy")
                ElseIf vDataType = ClsMasterKompetensi.EnumDataType.Text Then
                    Return vPar
                ElseIf vDataType = ClsMasterKompetensi.EnumDataType.YN Then
                    If vPar = "1" Then
                        Return "Yes"
                    Else
                        Return "No"
                    End If
                ElseIf vDataType = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                    If vPar = ClsMasterKompetensi.Aphabetic.A Then
                        Return "A"
                    ElseIf vPar = ClsMasterKompetensi.Aphabetic.B Then
                        Return "B"
                    ElseIf vPar = ClsMasterKompetensi.Aphabetic.C Then
                        Return "C"
                    ElseIf vPar = ClsMasterKompetensi.Aphabetic.D Then
                        Return "D"
                    ElseIf vPar = ClsMasterKompetensi.Aphabetic.E Then
                        Return "E"
                    ElseIf vPar = ClsMasterKompetensi.Aphabetic.F Then
                        Return "F"
                    End If
                End If
            Else
                Return ""
            End If


            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetValue_CustomFormatted(ByVal vPrimary As Double, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vDataTipe As ClsMasterKompetensi.EnumDataType) As String
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Dim Def_Prec As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetValue(vPrimary)) = False Then
                Def_Prec = Me.GetDefPrec(vPrimary)
                Return Me.SetFormatDataCustom(DTA.ScalarQuery_GetValue(vPrimary), vDataTipe, Def_Prec, vObjCD)
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

    Public Function GetJmlByDataCode(ByVal vDataCode As String) As Double
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByDataCode(vDataCode).HasValue Then
                Return DTA.ScalarQuery_GetJmlByDataCode(vDataCode).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_Gettimestamp(vPrimary).HasValue Then
                Return DTA.ScalarQuery_Gettimestamp(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vDataCode As String) As Double
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimary(vDataCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vDataCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetCreatedBy(vPrimary)
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

    Public Function GetDataCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If String.IsNullOrEmpty(DTA.ScalarQuery_GetJmlByDataCode(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetDataCode(vPrimary)
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

    ''' <summary>
    ''' data ini digunakan untuk mengecek kebenaran data yang diberikan
    ''' </summary>
    ''' <param name="vBaru"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vShowMsg As Boolean) As Boolean
        Try
            'tidak boleh ada kode yang sama
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mDataCode) = True Then
                    If vShowMsg = True Then
                        MsgBox("Code duplication, please change code data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.GetDataCode(Me.mPrimary)
                If CodeLama <> Me.mDataCode Then
                    If Me.ApakahSudahAda(Me.mDataCode) = True Then
                        If vShowMsg = True Then
                            MsgBox("Code duplication, please change code data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            'yang harus diisikan:
            '1. kode
            If Me.DataCode = "" Then
                If vShowMsg = True Then
                    MsgBox("Please specify code.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '2. description
            If Me.Description = "" Then
                If vShowMsg = True Then
                    MsgBox("Please specify description.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If


            'kalo mpe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterKompetensi, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntri.GetWaitTime Then
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
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True, True) = True Then
                If DTA.Insert(Me.mDataCode, Me.mDesc, Me.mDataType, Me.mDefaultPrecision, _
                Me.mValue, Me.mRemarks, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                    Me.mPrimary = Me.GetPrimary(Me.mDataCode)
                    Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

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

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            'ambil no antrian
            Me.ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterKompetensi, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntri.GetWaitTime Then
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
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(False, True) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                If DTA.Update(Me.mDataCode, Me.mDesc, Me.mDataType, _
                Me.mDefaultPrecision, Me.mValue, Me.mRemarks, _
                vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
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

    Public Overridable Function Hapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            'ambil no antrian
            Me.ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterKompetensi, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntri.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimary, vUserCode, _
                        vStrictDelete, vShowMsg)
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

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean) As Boolean
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUserCode, vStrictDelete, vShowMsg) = True Then
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

    Public Function ApakahSudahAda(ByVal vKode As String) As Boolean
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByDataCode(vKode).HasValue Then
                If DTA.ScalarQuery_GetJmlByDataCode(vKode).Value > 0 Then
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

    Public Function ApakahSudahDigunakan(ByVal vPrimary As Double, ByVal vShowMsg As Boolean) As Boolean
        Dim DTAJabatan As New DseMasterKompetensiTableAdapters.tblmaster_jobposition_item_kompetensiTableAdapter
        Dim DTAKaryawan As New DseMasterKompetensiTableAdapters.tblmaster_karyawan_item_kompetensiTableAdapter
        Dim DTART As New DseMasterKompetensiTableAdapters.tbltrans_rencana_training_item_kompetensiTableAdapter
        Dim dummy As New Nullable(Of Long)
        Dim TipeParameter As ClsMasterSalaryFormula_item.EnumTipeConstraitVariable = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi
        Dim DTA_OtherDeduction_Variable As New DseMasterKompetensiTableAdapters.tblmaster_salary_formula_otherdeduction_variableTableAdapter
        Dim TBL_OtherDeduction_Variable As New DseMasterKompetensi.tblmaster_salary_formula_otherdeduction_variableDataTable
        Dim row_OtherDeduction_Variable As DseMasterKompetensi.tblmaster_salary_formula_otherdeduction_variableRow
        Dim DTA_OtherDeduction_Constraint As New DseMasterKompetensiTableAdapters.tblmaster_salary_formula_otherdeduction_constraintTableAdapter
        Dim TBL_OtherDeduction_Constraint As New DseMasterKompetensi.tblmaster_salary_formula_otherdeduction_constraintDataTable
        Dim row_OtherDeduction_Constraint As DseMasterKompetensi.tblmaster_salary_formula_otherdeduction_constraintRow
        Dim ListException As New List(Of String)
        Dim DTA_SalaryFormula_Variable As New DseMasterKompetensiTableAdapters.tblmaster_salary_formula_variableTableAdapter
        Dim TBL_SalaryFormula_Variable As New DseMasterKompetensi.tblmaster_salary_formula_variableDataTable
        Dim row_SalaryFormula_Variable As DseMasterKompetensi.tblmaster_salary_formula_variableRow
        Dim DTA_SalaryFormula_Constraint As New DseMasterKompetensiTableAdapters.tblmaster_salary_formula_constraintTableAdapter
        Dim TBL_SalaryFormula_Constraint As New DseMasterKompetensi.tblmaster_salary_formula_constraintDataTable
        Dim row_SalaryFormula_Constraint As DseMasterKompetensi.tblmaster_salary_formula_constraintRow

        Dim DummyStr As String = ""
        Dim DummyInt As Integer = 0
        Dim DummyStr2 As String = ""

        Try
            DTAJabatan.Connection.ConnectionString = Core_Module.GetConstring
            DTAKaryawan.Connection.ConnectionString = DTAJabatan.Connection.ConnectionString
            DTART.Connection.ConnectionString = DTAJabatan.Connection.ConnectionString
            DTA_OtherDeduction_Variable.Connection.ConnectionString = DTAJabatan.Connection.ConnectionString
            DTA_OtherDeduction_Constraint.Connection.ConnectionString = DTAJabatan.Connection.ConnectionString
            DTA_SalaryFormula_Constraint.Connection.ConnectionString = DTAJabatan.Connection.ConnectionString
            DTA_SalaryFormula_Variable.Connection.ConnectionString = DTAJabatan.Connection.ConnectionString

            '1.tidak boleh dihapus bila sudah ada yang merefer:
            '   a.master jabatan, master karyawan, rencana training
            dummy = DTAJabatan.ScalarQuery_GetJml(vPrimary)
            If dummy.HasValue = True Then
                If dummy.Value > 0 Then
                    If vShowMsg = True Then
                        MsgBox("There are job positions for this data.", MsgBoxStyle.Information, "Information")
                    End If
                    Return True
                End If
            End If
            dummy = DTAKaryawan.ScalarQuery_GetJml(vPrimary)
            If dummy.HasValue = True Then
                If dummy.Value > 0 Then
                    If vShowMsg = True Then
                        MsgBox("There are employees for this data.", MsgBoxStyle.Information, "Information")
                    End If
                    Return True
                End If
            End If
            dummy = DTART.ScalarQuery_GetJml(vPrimary)
            If dummy.HasValue = True Then
                If dummy.Value > 0 Then
                    If vShowMsg = True Then
                        MsgBox("There are training plans for this data.", MsgBoxStyle.Information, "Information")
                    End If
                    Return True
                End If
            End If

            'terkait dengan salary formula atau other deduction
            'other deduction - variable
            DTA_OtherDeduction_Variable.Fill(TBL_OtherDeduction_Variable, vPrimary, CInt(TipeParameter))
            If TBL_OtherDeduction_Variable.Rows.Count > 0 Then
                ListException.Clear()
                For Each row_OtherDeduction_Variable In TBL_OtherDeduction_Variable.Rows
                    If ListException.Contains(row_OtherDeduction_Variable.Code + " - " + row_OtherDeduction_Variable.Description) = False Then
                        ListException.Add(row_OtherDeduction_Variable.Code + " - " + row_OtherDeduction_Variable.Description)
                    End If
                Next
                If ListException.Count > 0 Then
                    DummyStr = ""
                    DummyInt = 0
                    For Each DummyStr2 In ListException
                        DummyInt += 1
                        If DummyStr = "" Then
                            DummyStr = DummyInt.ToString + ". " + DummyStr2
                        Else
                            DummyStr = DummyStr + vbCrLf + DummyInt.ToString + ". " + DummyStr2
                        End If
                    Next
                    MsgBox("There are other deduction using this data: " + vbCrLf + DummyStr, MsgBoxStyle.Information, "Information")
                    Return True
                End If
            End If
            DTA_OtherDeduction_Constraint.Fill(TBL_OtherDeduction_Constraint, vPrimary, CInt(TipeParameter))
            If TBL_OtherDeduction_Constraint.Rows.Count > 0 Then
                ListException.Clear()
                For Each row_OtherDeduction_Constraint In TBL_OtherDeduction_Constraint.Rows
                    If ListException.Contains(row_OtherDeduction_Constraint.Code + " - " + row_OtherDeduction_Constraint.Description) = False Then
                        ListException.Add(row_OtherDeduction_Constraint.Code + " - " + row_OtherDeduction_Constraint.Description)
                    End If
                Next
                If ListException.Count > 0 Then
                    DummyStr = ""
                    DummyInt = 0
                    For Each DummyStr2 In ListException
                        DummyInt += 1
                        If DummyStr = "" Then
                            DummyStr = DummyInt.ToString + ". " + DummyStr2
                        Else
                            DummyStr = DummyStr + vbCrLf + DummyInt.ToString + ". " + DummyStr2
                        End If
                    Next
                    MsgBox("There are other deduction using this data: " + vbCrLf + DummyStr, MsgBoxStyle.Information, "Information")
                    Return True
                End If
            End If

            DTA_SalaryFormula_Variable.Fill(TBL_SalaryFormula_Variable, vPrimary, CInt(TipeParameter))
            If TBL_SalaryFormula_Variable.Rows.Count > 0 Then
                ListException.Clear()
                For Each row_SalaryFormula_Variable In TBL_SalaryFormula_Variable.Rows
                    If ListException.Contains(row_SalaryFormula_Variable.code + " - " + row_SalaryFormula_Variable.desc) = False Then
                        ListException.Add(row_SalaryFormula_Variable.code + " - " + row_SalaryFormula_Variable.desc)
                    End If
                Next
                If ListException.Count > 0 Then
                    DummyStr = ""
                    DummyInt = 0
                    For Each DummyStr2 In ListException
                        DummyInt += 1
                        If DummyStr = "" Then
                            DummyStr = DummyInt.ToString + ". " + DummyStr2
                        Else
                            DummyStr = DummyStr + vbCrLf + DummyInt.ToString + ". " + DummyStr2
                        End If
                    Next
                    MsgBox("There salary formula using this data: " + vbCrLf + DummyStr, MsgBoxStyle.Information, "Information")
                    Return True
                End If
            End If
            DTA_SalaryFormula_Constraint.Fill(TBL_SalaryFormula_Constraint, vPrimary, CInt(TipeParameter))
            If TBL_SalaryFormula_Constraint.Rows.Count > 0 Then
                ListException.Clear()
                For Each row_SalaryFormula_Constraint In TBL_SalaryFormula_Constraint.Rows
                    If ListException.Contains(row_SalaryFormula_Constraint.code + " - " + row_SalaryFormula_Constraint.desc) = False Then
                        ListException.Add(row_SalaryFormula_Constraint.code + " - " + row_SalaryFormula_Constraint.desc)
                    End If
                Next
                If ListException.Count > 0 Then
                    DummyStr = ""
                    DummyInt = 0
                    For Each DummyStr2 In ListException
                        DummyInt += 1
                        If DummyStr = "" Then
                            DummyStr = DummyInt.ToString + ". " + DummyStr2
                        Else
                            DummyStr = DummyStr + vbCrLf + DummyInt.ToString + ". " + DummyStr2
                        End If
                    Next
                    MsgBox("There salary formula using this data: " + vbCrLf + DummyStr, MsgBoxStyle.Information, "Information")
                    Return True
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        Finally
            DTAJabatan.Dispose()
            DTART.Dispose()
            DTAKaryawan.Dispose()
            DTA_OtherDeduction_Variable.Dispose()
            TBL_OtherDeduction_Variable.Dispose()
            DTA_OtherDeduction_Constraint.Dispose()
            TBL_OtherDeduction_Constraint.Dispose()
            DTA_SalaryFormula_Variable.Dispose()
            TBL_SalaryFormula_Variable.Dispose()
            DTA_SalaryFormula_Constraint.Dispose()
            TBL_SalaryFormula_Constraint.Dispose()
            ListException = Nothing
        End Try
    End Function

    Public Function ApakahValid_Aphabetic(ByVal vNilai As String) As Boolean
        Try
            If vNilai = "A" Or vNilai = "B" Or vNilai = "C" _
            Or vNilai = "D" Or vNilai = "E" Or vNilai = "F" Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function ApakahTipeDataTepat(ByVal vDGV As DataGridView, ByVal vIndexPrimary As Integer, ByVal vIndexTipeData As Integer, _
    ByVal vIndexCode As String, ByVal vShowMsg As Boolean) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAKompetensi As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Dim TBLKompetensi As New DseMasterKompetensi.tblmaster_kompetensiDataTable
        Dim PrimaryMin, PrimaryMax As Double
        Dim DummyBool As Boolean = False

        Try
            DTAKompetensi.Connection.ConnectionString = Core_Module.GetConstring

            DummyBool = vDGV.AllowUserToAddRows
            vDGV.AllowUserToAddRows = False
            If vDGV.Rows.Count = 0 Then
                vDGV.AllowUserToAddRows = DummyBool
                Return False
            End If

            PrimaryMax = Double.MinValue
            PrimaryMin = Double.MaxValue

            For Each rowDGV In vDGV.Rows
                If rowDGV.Cells(vIndexPrimary).Value < PrimaryMin Then
                    PrimaryMin = rowDGV.Cells(vIndexPrimary).Value
                End If
                If rowDGV.Cells(vIndexPrimary).Value > PrimaryMax Then
                    PrimaryMax = rowDGV.Cells(vIndexPrimary).Value
                End If
            Next
            DTAKompetensi.FillBy_PrimaryMinMax(TBLKompetensi, PrimaryMin, PrimaryMax)

            For Each rowDGV In vDGV.Rows
                TBLKompetensi.DefaultView.RowFilter = "Primary = " + rowDGV.Cells(vIndexPrimary).Value.ToString + _
                    " AND [Data Type] = " + rowDGV.Cells(vIndexTipeData).Value.ToString

                If TBLKompetensi.DefaultView.Count <= 0 Then
                    If vShowMsg = True Then
                        MsgBox("Inconsistent data type for: " + rowDGV.Cells(CInt(vIndexCode)).Value, MsgBoxStyle.Information, "Information")
                    End If
                    vDGV.AllowUserToAddRows = DummyBool
                    Return False
                End If
            Next

            vDGV.AllowUserToAddRows = DummyBool
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAKompetensi.Dispose()
            TBLKompetensi.Dispose()
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal vStrictEdit As Boolean, _
    ByVal vShowMsg As Boolean) As Boolean
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Dim pCreatedBy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            '1. tidak boleh hapus jika melanggar aturan strict edit
            If vStrictEdit = True Then
                If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                Else
                    pCreatedBy = ""
                End If

                If pCreatedBy <> vUserCode Then
                    If vShowMsg = True Then
                        MsgBox("User can edit this data: " + pCreatedBy, MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal vStrictDelete As Boolean, _
    ByVal vShowMsg As Boolean) As Boolean
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Dim pCreatedBy As String = ""
        Dim DTAJabatan As New DseMasterKompetensiTableAdapters.tblmaster_jobposition_item_kompetensiTableAdapter
        Dim DTAKaryawan As New DseMasterKompetensiTableAdapters.tblmaster_karyawan_item_kompetensiTableAdapter
        Dim DTART As New DseMasterKompetensiTableAdapters.tbltrans_rencana_training_item_kompetensiTableAdapter
        Dim dummy As New Nullable(Of Long)
        Dim TipeParameter As ClsMasterSalaryFormula_item.EnumTipeConstraitVariable = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi
        Dim DTA_OtherDeduction_Variable As New DseMasterKompetensiTableAdapters.tblmaster_salary_formula_otherdeduction_variableTableAdapter
        Dim TBL_OtherDeduction_Variable As New DseMasterKompetensi.tblmaster_salary_formula_otherdeduction_variableDataTable
        Dim row_OtherDeduction_Variable As DseMasterKompetensi.tblmaster_salary_formula_otherdeduction_variableRow
        Dim DTA_OtherDeduction_Constraint As New DseMasterKompetensiTableAdapters.tblmaster_salary_formula_otherdeduction_constraintTableAdapter
        Dim TBL_OtherDeduction_Constraint As New DseMasterKompetensi.tblmaster_salary_formula_otherdeduction_constraintDataTable
        Dim row_OtherDeduction_Constraint As DseMasterKompetensi.tblmaster_salary_formula_otherdeduction_constraintRow
        Dim ListException As New List(Of String)
        Dim DTA_SalaryFormula_Variable As New DseMasterKompetensiTableAdapters.tblmaster_salary_formula_variableTableAdapter
        Dim TBL_SalaryFormula_Variable As New DseMasterKompetensi.tblmaster_salary_formula_variableDataTable
        Dim row_SalaryFormula_Variable As DseMasterKompetensi.tblmaster_salary_formula_variableRow
        Dim DTA_SalaryFormula_Constraint As New DseMasterKompetensiTableAdapters.tblmaster_salary_formula_constraintTableAdapter
        Dim TBL_SalaryFormula_Constraint As New DseMasterKompetensi.tblmaster_salary_formula_constraintDataTable
        Dim row_SalaryFormula_Constraint As DseMasterKompetensi.tblmaster_salary_formula_constraintRow
        Dim DummyStr As String = ""
        Dim DummyInt As Integer = 0
        Dim DummyStr2 As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAJabatan.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAKaryawan.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTART.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_OtherDeduction_Constraint.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_OtherDeduction_Variable.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_SalaryFormula_Constraint.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_SalaryFormula_Variable.Connection.ConnectionString = DTA.Connection.ConnectionString

            '1.tidak boleh dihapus bila sudah ada yang merefer:
            '   a.master jabatan, master karyawan, rencana training
            dummy = DTAJabatan.ScalarQuery_GetJml(vPrimary)
            If dummy.HasValue = True Then
                If dummy.Value > 0 Then
                    MsgBox("There are job positions for this data, you cannot delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            dummy = DTAKaryawan.ScalarQuery_GetJml(vPrimary)
            If dummy.HasValue = True Then
                If dummy.Value > 0 Then
                    MsgBox("There are employees for this data, you cannot delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            dummy = DTART.ScalarQuery_GetJml(vPrimary)
            If dummy.HasValue = True Then
                If dummy.Value > 0 Then
                    MsgBox("There are training plans for this data, you cannot delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'terkait dengan salary formula atau other deduction
            'terkait dengan salary formula atau other deduction
            'other deduction - variable
            DTA_OtherDeduction_Variable.Fill(TBL_OtherDeduction_Variable, vPrimary, CInt(TipeParameter))
            If TBL_OtherDeduction_Variable.Rows.Count > 0 Then
                ListException.Clear()
                For Each row_OtherDeduction_Variable In TBL_OtherDeduction_Variable.Rows
                    If ListException.Contains(row_OtherDeduction_Variable.Code + " - " + row_OtherDeduction_Variable.Description) = False Then
                        ListException.Add(row_OtherDeduction_Variable.Code + " - " + row_OtherDeduction_Variable.Description)
                    End If
                Next
                If ListException.Count > 0 Then
                    DummyStr = ""
                    DummyInt = 0
                    For Each DummyStr2 In ListException
                        DummyInt += 1
                        If DummyStr = "" Then
                            DummyStr = DummyInt.ToString + ". " + DummyStr2
                        Else
                            DummyStr = DummyStr + vbCrLf + DummyInt.ToString + ". " + DummyStr2
                        End If
                    Next
                    MsgBox("There are other deduction using this data: " + vbCrLf + DummyStr + vbCrLf + "You can't delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            DTA_OtherDeduction_Constraint.Fill(TBL_OtherDeduction_Constraint, vPrimary, CInt(TipeParameter))
            If TBL_OtherDeduction_Constraint.Rows.Count > 0 Then
                ListException.Clear()
                For Each row_OtherDeduction_Constraint In TBL_OtherDeduction_Constraint.Rows
                    If ListException.Contains(row_OtherDeduction_Constraint.Code + " - " + row_OtherDeduction_Constraint.Description) = False Then
                        ListException.Add(row_OtherDeduction_Constraint.Code + " - " + row_OtherDeduction_Constraint.Description)
                    End If
                Next
                If ListException.Count > 0 Then
                    DummyStr = ""
                    DummyInt = 0
                    For Each DummyStr2 In ListException
                        DummyInt += 1
                        If DummyStr = "" Then
                            DummyStr = DummyInt.ToString + ". " + DummyStr2
                        Else
                            DummyStr = DummyStr + vbCrLf + DummyInt.ToString + ". " + DummyStr2
                        End If
                    Next
                    MsgBox("There are other deduction using this data: " + vbCrLf + DummyStr + vbCrLf + "You can't delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            DTA_SalaryFormula_Variable.Fill(TBL_SalaryFormula_Variable, vPrimary, CInt(TipeParameter))
            If TBL_SalaryFormula_Variable.Rows.Count > 0 Then
                ListException.Clear()
                For Each row_SalaryFormula_Variable In TBL_SalaryFormula_Variable.Rows
                    If ListException.Contains(row_SalaryFormula_Variable.code + " - " + row_SalaryFormula_Variable.desc) = False Then
                        ListException.Add(row_SalaryFormula_Variable.code + " - " + row_SalaryFormula_Variable.desc)
                    End If
                Next
                If ListException.Count > 0 Then
                    DummyStr = ""
                    DummyInt = 0
                    For Each DummyStr2 In ListException
                        DummyInt += 1
                        If DummyStr = "" Then
                            DummyStr = DummyInt.ToString + ". " + DummyStr2
                        Else
                            DummyStr = DummyStr + vbCrLf + DummyInt.ToString + ". " + DummyStr2
                        End If
                    Next
                    MsgBox("There salary formula using this data: " + vbCrLf + DummyStr + vbCrLf + "You can't delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            DTA_SalaryFormula_Constraint.Fill(TBL_SalaryFormula_Constraint, vPrimary, CInt(TipeParameter))
            If TBL_SalaryFormula_Constraint.Rows.Count > 0 Then
                ListException.Clear()
                For Each row_SalaryFormula_Constraint In TBL_SalaryFormula_Constraint.Rows
                    If ListException.Contains(row_SalaryFormula_Constraint.code + " - " + row_SalaryFormula_Constraint.desc) = False Then
                        ListException.Add(row_SalaryFormula_Constraint.code + " - " + row_SalaryFormula_Constraint.desc)
                    End If
                Next
                If ListException.Count > 0 Then
                    DummyStr = ""
                    DummyInt = 0
                    For Each DummyStr2 In ListException
                        DummyInt += 1
                        If DummyStr = "" Then
                            DummyStr = DummyInt.ToString + ". " + DummyStr2
                        Else
                            DummyStr = DummyStr + vbCrLf + DummyInt.ToString + ". " + DummyStr2
                        End If
                    Next
                    MsgBox("There salary formula using this data: " + vbCrLf + DummyStr + vbCrLf + "You can't delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '2. tidak boleh hapus jika melanggar aturan strict edit
            If vStrictDelete = True Then
                If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                Else
                    pCreatedBy = ""
                End If

                If pCreatedBy <> vUserCode Then
                    If vShowMsg = True Then
                        MsgBox("User can delete this data: " + pCreatedBy, MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAJabatan.Dispose()
            DTART.Dispose()
            DTAKaryawan.Dispose()
            DTA_OtherDeduction_Variable.Dispose()
            TBL_OtherDeduction_Variable.Dispose()
            DTA_OtherDeduction_Constraint.Dispose()
            TBL_OtherDeduction_Constraint.Dispose()
            DTA_SalaryFormula_Variable.Dispose()
            TBL_SalaryFormula_Variable.Dispose()
            DTA_SalaryFormula_Constraint.Dispose()
            TBL_SalaryFormula_Constraint.Dispose()
            ListException = Nothing
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterKompetensi.tblmaster_kompetensi1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal vNumberOnly As Boolean) As Boolean
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensi1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date = Now
        Dim TipeKompetensi As ClsMasterKompetensi.EnumDataType = EnumDataType.Number

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If vNumberOnly = True Then
                If SearchCriteria = EnumSearchCriteria.ByCode Then
                    QueryResult = DTA.FillBy_DataCodeLIKE2(TBL, Criteria + "%", CInt(TipeKompetensi))
                ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                    QueryResult = DTA.FillBy_DescriptionLIKE2(TBL, Criteria + "%", CInt(TipeKompetensi))
                ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                    QueryResult = DTA.FillBy_KeteranganLIKE2(TBL, Criteria + "%", CInt(TipeKompetensi))
                End If
            Else
                If SearchCriteria = EnumSearchCriteria.ByCode Then
                    QueryResult = DTA.FillBy_DataCodeLIKE(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                    QueryResult = DTA.FillBy_DescriptionLIKE(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                    QueryResult = DTA.FillBy_KeteranganLIKE(TBL, Criteria + "%")
                End If
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
                If TS.Ticks > TimeSpan.TicksPerSecond Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Seconds) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, _
                    0, MidpointRounding.AwayFromZero)) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If

            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function CariDetail(ByVal Primary As Double) As Boolean
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensiTableAdapter
        Dim TBL As New DseMasterKompetensi.tblmaster_kompetensiDataTable
        Dim row As DseMasterKompetensi.tblmaster_kompetensiRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, Primary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mDataCode = .Data_Code
                    Me.mDesc = .Description
                    Me.mDataType = .Data_Type
                    Me.mDefaultPrecision = .Default_Precision
                    Me.mValue = .Value
                    Me.mRemarks = .Remarks

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
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As EnumSearchCriteria, ByVal vNumberOnly As Boolean) As Double
        Dim DTA As New DseMasterKompetensiTableAdapters.tblmaster_kompetensi1TableAdapter
        Dim TipeKompetensi As EnumDataType = EnumDataType.Number

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vNumberOnly = True Then
                If vSearchBy = EnumSearchCriteria.ByCode Then
                    Return DTA.ScalarQuery_GetJmlByDataCodeLIKE2(vKriteria + "%", CInt(TipeKompetensi)).Value
                ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                    Return DTA.ScalarQuery_GetJmlByDescLIKE2(vKriteria + "%", CInt(TipeKompetensi)).Value
                ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                    Return DTA.ScalarQuery_GetJmlByRemarksLIKE2(vKriteria + "%", CInt(TipeKompetensi)).Value
                End If
            Else
                If vSearchBy = EnumSearchCriteria.ByCode Then
                    Return DTA.ScalarQuery_GetJmlByDataCodeLIKE(vKriteria + "%").Value
                ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                    Return DTA.ScalarQuery_GetJmlByDescLIKE(vKriteria + "%").Value
                ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                    Return DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%").Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
