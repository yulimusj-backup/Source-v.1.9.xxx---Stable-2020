Option Explicit On
Option Strict Off
Public Class ClsTransaksiPR_Cancel
    '''member pada form sendiri

    Private mPrimary As Double = -1
    Private mPrimaryPR As Double = -1
    Private mNoTransaksi As String = ""
    Private mTanggal As Date = Now
    Private mUrutan As Double = 0
    Private mReason As String = ""
    Private mRemarks As String = ""
    Private mCreatedBy As String = ""
    Private mCreatedDate As Date = Now
    Private mModifiedBy As String = ""
    Private mModifiedDate As Date = Now
    Private mTimestamp As Date = Now

    Dim ObjRefParty As New Core_Module.ClsMasterRefParty
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjPRCancel_Item As New ClsTransaksiPR_Cancel_Item
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjUser As New Core_Module.ClsUser
    Dim objPR As New ClsTransaksiPR

    '''khusus untuk antrian
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianPurchaseRequest
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

    Public Property PrimaryPR() As Double
        Get
            Return mPrimaryPR
        End Get
        Set(ByVal value As Double)
            mPrimaryPR = value
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

    Public Property Reason() As String
        Get
            Return mReason
        End Get
        Set(ByVal value As String)
            mReason = value
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

    Public Property TimeStamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
        End Set
    End Property

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancelTableAdapter
        Dim Retval As New Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetPrimary(vNoTransaksi)
            If Retval.HasValue = True Then
                If Retval.Value > -1 Then
                    Return Retval
                Else
                    Return -1
                End If
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryPR(ByVal vPrimaryPRCancel As Double) As Double
        Dim DTA As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancelTableAdapter
        Dim Retval As New Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetPrimaryPR(vPrimaryPRCancel)
            If Retval.HasValue = True Then
                If Retval.Value > -1 Then
                    Return Retval
                Else
                    Return -1
                End If
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancelTableAdapter
        Dim Retval As New Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            If Retval.HasValue Then
                Return Retval
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetQty(ByVal vPrimaryPR_Cancel As Double, ByVal vPrimaryCode As Double, ByVal vPrimaryPR As Double) As Double
        Dim DTA As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancel_itemTableAdapter
        Dim Qty As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Qty = DTA.ScalarQuery_GetQty(vPrimaryPR_Cancel, vPrimaryCode, vPrimaryPR)
            If Qty.HasValue = True Then
                Return Qty.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancelTableAdapter
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetNoTransaksi(vPrimary)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByNoTransaksi(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancelTableAdapter
        Dim Retval As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(vNoTransaksi)
            If Retval.HasValue = True Then
                If String.IsNullOrEmpty(Retval) = False Then
                    Return Retval
                Else
                    Return 0
                End If
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
           EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
                                     Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Double
        Dim DTA As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancelTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByNoTransaksi Then
                RetVal = DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByTanggalTransaksi Then
                RetVal = DTA.ScalarQuery_GetJmlByTanggalSampai(vTanggalDari, vTanggalSampai)
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByRequestBy Then
                RetVal = DTA.ScalarQuery_GetJmlByReqBy(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByReason Then
                RetVal = DTA.ScalarQuery_GetJmlByReasonLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If

            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                RetVal = DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByProductCodeNo Then
                RetVal = DTA.ScalarQuery_GetJmlByCodeNo(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByProductDesc Then
                RetVal = DTA.ScalarQuery_GetJmlByProductDesc(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
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

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancelTableAdapter
        Dim TanggalAwal, TanggalAkhir As Date
        Dim NoUrut As String
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TanggalAwal = ObjCD.GetFirstDateofTheMonth(vTanggalTransaksi)
            TanggalAkhir = ObjCD.GetLasDateOfTheMonth(vTanggalTransaksi)
            If DTA.ScalarQuery_GetMaxUrutan(TanggalAwal, TanggalAkhir).HasValue Then
                NoUrut = CStr(DTA.ScalarQuery_GetMaxUrutan(TanggalAwal, TanggalAkhir).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If
            If IsNumeric(NoUrut) Then
                Me.mUrutan = CDbl(NoUrut)
            End If
            While NoUrut.Length < 6
                NoUrut = "0" + NoUrut
            End While
            Return NoUrut + "/" + "PRC" + "/" + CStr(Format(vTanggalTransaksi, "MMM").ToUpper) + "/" + CStr(Format(vTanggalTransaksi, "yyyy"))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahQtyBenar(ByVal vPrimary As Double, ByVal vPrimaryPR As Double, _
                                   ByVal vDGV As DataGridView, ByVal vFrm As FrmTransaksiPR_Cancel) As Boolean

        Dim vPrimaryCodeNo As Double = -1
        Dim DefPrec As Decimal = 0
        Dim DGVQty As Decimal = 0
        Dim OB As Decimal = 0
        Dim ItemDB As Decimal = 0
        Dim Maximumnya As Decimal = 0
        Dim OB_StlhEdit As Decimal = 0
        Dim vUnit As String = ""
        Dim vCode As String = ""
        Dim DaftarMin As String = ""
        Try
            If vDGV.RowCount > 0 Then
                For i As Integer = 0 To vDGV.RowCount - 1
                    vPrimaryCodeNo = Me.ObjProduct.GetPrimary(vFrm.DgvItem(vFrm.ColItem_CodeNo.Index, i).Value)
                    vCode = Me.ObjProduct.GetProductCode(vPrimaryCodeNo)
                    vUnit = Me.ObjProduct.GetUnitCode(vPrimaryCodeNo)
                    DefPrec = Me.ObjProduct.GetAngkaDecimal(vPrimaryCodeNo)
                    DGVQty = vFrm.DgvItem.Item(vFrm.ColItem_Qty.Index, i).Value
                    OB = Me.objPR.GetOB(vPrimaryPR, vPrimaryCodeNo)
                    ItemDB = Me.GetQty(vPrimary, vPrimaryCodeNo, vPrimaryPR)

                    OB_StlhEdit = OB + ItemDB - DGVQty
                    If OB_StlhEdit < 0 Then
                        Maximumnya = DGVQty + OB_StlhEdit
                        MsgBox("Code No : " & vCode & " " & "maximum value you can enter is" & " " & Me.ObjCD.DigitGrouping(Maximumnya, DefPrec) & _
                        " " & vUnit.ToLower & " ", MsgBoxStyle.Information, "Information")
                        vFrm.DgvItem.ClearSelection()
                        vFrm.DgvItem.Rows(i).Selected = True
                        'vFrm.DgvItem(vFrm.ColItem_Qty.Index, i).Value = Me.ObjCD.DigitGrouping(Maximumnya, DefPrec)
                        ' vFrm.DgvItem(vFrm.ColItem_Qty.Index, i).Selected = True
                        Return False
                    End If

                    '''digit grouping
                    vFrm.DgvItem(vFrm.ColItem_Qty.Index, i).Value = _
                    Me.ObjCD.DigitGrouping(vFrm.DgvItem(vFrm.ColItem_Qty.Index, i).Value, DefPrec)
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    '''mengecek tanggal di pickup order tidak lebih kecil dengan tanggal pembuatan prnya sendiri
    Public Function ApakahTanggalSudahBenar(ByVal vPrimaryPR As Double, ByVal vTglDiPickupOrder As Date) As Boolean
        Try
            Dim Tgl_Di_Purchase_Request As Date = objPR.GetTanggalTransaksi(vPrimaryPR)
            If vTglDiPickupOrder.Date < Tgl_Di_Purchase_Request.Date = True Then
                MsgBox("Dating in purchase request cancel must be > or same with date in purchase request", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vPrimary As Double, _
                                    ByVal vPrimaryPR As Double, ByVal vDGV As DataGridView, _
                                    ByVal vFrm As FrmTransaksiPR_Cancel) As Boolean
        Dim ObjPRItem As New ClsTransaksiPR_Item
        Try
            '''cek notransaksi, blok jika sama
            vDGV.AllowUserToAddRows = False
            If vBaru = True Then
                'jika sudah ada, maka ganti no transaksi
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggal)
                End If
            Else
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    If Me.mPrimary <> Me.GetPrimary(Me.mNoTransaksi) Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggal)
                    End If
                End If
            End If

            If vDGV.RowCount > 0 Then
                If vBaru = True Then
                    vFrm.KondisiPRNumber(True)
                End If
                If Me.ApakahQtyBenar(vPrimary, vPrimaryPR, vDGV, vFrm) = False Then
                    Return False
                End If
            End If

            '''Item tidak boleh kosong 
            If vDGV.Rows.Count = 1 Then
                For i As Integer = 0 To vDGV.RowCount - 1
                    If vDGV.Item(vFrm.ColItem_CodeNo.Index, i).Value = "" Then
                        MsgBox("Please input product code", MsgBoxStyle.Information, "Information ")
                        Return False
                    End If
                Next
            ElseIf vDGV.RowCount = 0 Then
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'vDGV.AllowUserToAddRows = True
        End Try
    End Function

    '''cek di Transaksi sendiri
    Public Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Try
            If Me.GetJmlByNoTransaksi(vNoTransaksi) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, _
                                        ByVal vSalesStaff As String, ByVal Frm As FrmTransaksiPR_Cancel, _
                                        Optional ByVal vTampilkanPesan As Boolean = True) As Boolean

        Dim DTA As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancelTableAdapter
        Dim Created_By As String = DTA.ScalarQuery_GetCreatedBy(vPrimary)
        Try
            If Created_By <> vSalesStaff Then
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

    Public Function ApakahBolehEdit(ByVal frm As FrmTransaksiPR_Cancel, ByVal vPrimary As Double, ByVal vSalesStaff As String, Optional ByVal vTampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancelTableAdapter
        Dim Created_By As String = DTA.ScalarQuery_GetCreatedBy(vPrimary)
        Try
            If Created_By <> vSalesStaff Then
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

    Public Function ApakahBolehHapusItem() As Boolean

    End Function

    Public Function ApakahBolehEditItem() As Boolean

    End Function

    Public Function SimpanBaru1(ByVal vDGVItem As DataGridView, _
                                    ByVal vUserCode As String, ByVal vFrm As FrmTransaksiPR_Cancel, _
                                    ByVal vBaru As Boolean, ByVal vPrimary As Double, _
                                    ByVal vPrimaryPR As Double) As Boolean

        Dim DTA As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancelTableAdapter
        Try

            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(vBaru, vPrimary, vPrimaryPR, vDGVItem, vFrm) = False Then Exit Function

            Me.mUrutan = CDbl(Me.mNoTransaksi.Substring(0, 6))
            If DTA.Insert(Me.mPrimaryPR, _
                          Me.mNoTransaksi, Me.mTanggal, _
                          Me.mUrutan, Me.mReason, Me.mRemarks, _
                          vUserCode, Now, vUserCode, Now, Now) > 0 Then

                Me.mPrimary = Me.GetPrimary(Me.mNoTransaksi)
                Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode
                'simpan data item
                Me.ObjPRCancel_Item.SimpanItem(Me.mPrimary, vFrm.DgvItem, vFrm)
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

    Public Overridable Function SimpanPerubahan1(ByVal vDGVItem As DataGridView, _
                                    ByVal vUserCode As String, ByVal vFrm As FrmTransaksiPR_Cancel, _
                                    ByVal vBaru As Boolean, ByVal vPrimary As Double, _
                                    ByVal vPrimaryPR As Double) As Boolean

        Dim PrimaryPRLama As Nullable(Of Double)
        Dim DTA As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancelTableAdapter
        Dim ObjPR As New ClsTransaksiPR

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(vBaru, vPrimary, vPrimaryPR, vDGVItem, vFrm) = False Then Exit Function
            Me.mUrutan = CDbl(Me.mNoTransaksi.Substring(0, 6))
            PrimaryPRLama = DTA.ScalarQuery_GetPrimaryPR(mPrimary)

            If DTA.Update(Me.mPrimaryPR, _
                          Me.mNoTransaksi, Me.mTanggal, _
                          Me.mUrutan, Me.mReason, Me.mRemarks, vUserCode, Now, Now, mPrimary) > 0 Then

                'update data yang berubah setelah simpan perubahan
                Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode

                'hapus data item lama
                ObjPRCancel_Item.HapusItem(Me.mPrimary, vFrm)
                If CDbl(PrimaryPRLama.Value) <> CDbl(vFrm.TxtPrimaryPR.Text) Then
                    ObjPR.SetOrderBalance1(PrimaryPRLama.Value, ClsTransaksiPR.EnumReadyForPurchase.Yes)
                End If

                'simpan data item baru
                ObjPRCancel_Item.SimpanItem(Me.mPrimary, vFrm.DgvItem, vFrm)

                'nilai balikan
                Return True
            Else
                Return False
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanBaru(ByVal vDGVItem As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiPR_Cancel, _
                               ByVal vBaru As Boolean, ByVal vPrimary As Double, ByVal vPrimaryPR As Double) As Boolean

        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier.ToString, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Purchase_Request_Cancel_SimpanBaru, vUserCode)
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
                            retval = Me.SimpanBaru1(vDGVItem, vUserCode, vFrm, vBaru, vPrimary, vPrimaryPR)
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
            'Me.DTAInOUT.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vDGVItem As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiPR_Cancel, _
                               ByVal vBaru As Boolean, ByVal vPrimary As Double, ByVal vPrimaryPR As Double) As Boolean

        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Purchase_Request_Cancel_SimpanPerubahan, vUserCode)

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
                            retval = Me.SimpanPerubahan1(vDGVItem, vUserCode, vFrm, vBaru, vPrimary, vPrimaryPR)
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
            'ObjAntri = Nothing
            SW = Nothing
            'Me.DTAInOUT.Dispose()
        End Try
    End Function

    Public Overridable Function Hapus1(ByVal vPrimary As Double, ByVal frm As FrmTransaksiPR_Cancel, ByVal StrictSalesStaffDelete As String) As Boolean
        Dim DTA As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancelTableAdapter
        Dim PrimaryPR As Double = frm.TxtPrimaryPR.Text
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary, StrictSalesStaffDelete, frm, True) = True Then
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimary) > 0 Then
                    Me.objPR.SetOrderBalance1(PrimaryPR, ClsTransaksiPR.EnumReadyForPurchase.Yes)
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal usercode As String, _
                            ByVal frm As FrmTransaksiPR_Cancel) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Purchase_Request_Cancel_Hapus, usercode)

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
                            retval = Me.Hapus1(vPrimaryTransaksi, frm, usercode)
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
            'ObjAntri = Nothing
            'Me.DTAInOUT.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
                                   ByVal vObjCD As Core_Module.ClsCekData, ByVal vFrm As FrmTransaksiPR_Cancel) As Boolean
        Dim DTA As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancelTableAdapter
        Dim TBL As New DseTransaksiPR_Cancel.tbltrans_purchase_req_cancelDataTable
        Dim row As DseTransaksiPR_Cancel.tbltrans_purchase_req_cancelRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, CDec(vPrimary))
            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseTransaksiPR_Cancel.tbltrans_purchase_req_cancelRow)
                With row
                    Me.mPrimary = .primary
                    Me.PrimaryPR = .primary_pr
                    Me.mNoTransaksi = .no_transaksi
                    Me.mTanggal = .tanggal
                    Me.mUrutan = .urutan
                    Me.mReason = .reason
                    Me.mRemarks = .remarks
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With
                ObjPRCancel_Item.IsiDGVItem(vPrimary, vDGVItem, vFrm, vObjCD)
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
         ByVal DGV As DataGridView, ByVal TBL As DseTransaksiPR_Cancel.tbltrans_purchase_req_cancel1DataTable, _
         ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
         Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
         Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiPR_CancelTableAdapters.tbltrans_purchase_req_cancel1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByPRNumber Then
                QueryResult = DTA.FillBy_NoTransaksiPRLIKE(TBL, Criteria + "%")

            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TglDariTglSampai(TBL, vTanggalDari, vTanggalSampai)

            ElseIf SearchCriteria = EnumSearchCriteria.ByReason Then
                QueryResult = DTA.FillBy_ReasonLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRequestBy Then
                QueryResult = DTA.FillBy_ReqByLIKE(TBL, Criteria + "%")

            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, Criteria + "%")
            End If

            'update DGV
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).Visible = False
                DGV.Columns(2).HeaderText = "Document No"
                DGV.Columns(3).HeaderText = "PR Number"
                DGV.Columns(4).HeaderText = "Date"
                DGV.Columns(6).HeaderText = "Reason"
                DGV.Columns(7).HeaderText = "Remarks"

                DGV.Columns(4).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(2).MinimumWidth = 100
                DGV.Columns(3).MinimumWidth = 100
                DGV.Columns(4).MinimumWidth = 100
                DGV.Columns(5).MinimumWidth = 80
                DGV.Columns(6).MinimumWidth = 80
                DGV.Columns(7).MinimumWidth = 80
            End If
            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 0, MidpointRounding.AwayFromZero) + _
                    " miliseconds"
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
            TBL.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function IsiDGV2(ByVal vPrimaryPR As Double, ByVal vPrimaryPRCancel As Double, _
                            ByVal vDGV As DataGridView, ByVal vFrm As FrmTransaksiPR_Cancel, _
                            Optional ByVal vShowMsg As Boolean = True) As Boolean

        Dim DTA As New DseTransaksiPR_CancelTableAdapters.TblIsiDGVItemTableAdapter
        Dim TBL As New DseTransaksiPR_Cancel.TblIsiDGVItemDataTable
        Dim row As DseTransaksiPR_Cancel.TblIsiDGVItemRow
        Dim rowDGV As DataGridViewRow
        Dim ObjUnit As New Core_Module.ClsMasterUnit
        Dim PrimaryLama As New Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.Rows.Clear()
            PrimaryLama = Me.GetPrimaryPR(vPrimaryPRCancel)
            If PrimaryLama <> vFrm.TxtPrimaryPR.Text Then
                If DTA.FillBy_PrimaryPR(TBL, vPrimaryPR) > 0 Then
                    For Each row In TBL
                        If row.OB > 0 Then
                            vDGV.Rows.Add()
                            vDGV.AllowUserToAddRows = False
                            rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                            Dim DefPrec As Integer = ObjProduct.GetAngkaDecimal(row.primary_code_no)
                            With rowDGV
                                .Cells(vFrm.ColItem_No.Index).Value = ""
                                .Cells(vFrm.ColItem_PrimaryCode.Index).Value = row.primary_code_no
                                .Cells(vFrm.ColItem_CodeNo.Index).Value = ObjProduct.GetProductCode(row.primary_code_no)
                                .Cells(vFrm.ColItem_Desc.Index).Value = ObjProduct.GetDescription(row.primary_code_no)
                                .Cells(vFrm.ColItem_Unit.Index).Value = ObjUnit.GetUnitCode(row.primary_unit)
                                .Cells(vFrm.ColItem_Qty.Index).Value = ObjCD.DigitGrouping(row.OB, DefPrec)
                                .Cells(vFrm.ColItem_Remarks.Index).Value = ""
                            End With
                        End If
                    Next
                End If
            Else
                Me.ObjPRCancel_Item.IsiDGVItem(vPrimaryPRCancel, vDGV, vFrm, ObjCD)
            End If

            vDGV.AllowUserToAddRows = False
            '''setnumber
            ObjCD.SetNumber(vDGV)
            '''autosize dgv
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            '''nilai balikan true
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
        Return True
    End Function

    Public Function CetakPR_Cancel(ByVal vPrimaryPR_Cancel As Double, _
                            ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer) As Boolean
        Dim Frm As New FrmTransaksiPR_Cancel_Cetak
        Dim Report As New CRPLaporanPR_Cancel
        Dim DS As New DseTransaksiPR_Cancel

        Dim DTACetak_Header As New DseTransaksiPR_CancelTableAdapters.TBLCetak_HeaderTableAdapter
        Dim TBLCetak_Header As DseTransaksiPR_Cancel.TBLCetak_HeaderDataTable = DS.TBLCetak_Header
        Dim RowCetak_Header As DseTransaksiPR_Cancel.TBLCetak_HeaderRow

        Dim DTACetak_Item As New DseTransaksiPR_CancelTableAdapters.TBLCetak_ItemTableAdapter
        Dim TBLCetak_Item As DseTransaksiPR_Cancel.TBLCetak_ItemDataTable = DS.TBLCetak_Item
        Dim RowCetak_Item As DseTransaksiPR_Cancel.TBLCetak_ItemRow

        Dim TBLHeader As DseTransaksiPR_Cancel.TBLHeaderDataTable = DS.TBLHeader
        Dim TBLItem As DseTransaksiPR_Cancel.TblItemDataTable = DS.TblItem

        Dim rowHeader As DseTransaksiPR_Cancel.TBLHeaderRow
        Dim rowItem As DseTransaksiPR_Cancel.TblItemRow

        Dim DV1 As DataRowView
        Dim DV2 As DataRowView
        Try
            DTACetak_Header.Connection.ConnectionString = Core_Module.GetConstring
            DTACetak_Item.Connection.ConnectionString = DTACetak_Header.Connection.ConnectionString
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            TBLHeader.Rows.Clear()
            TBLItem.Rows.Clear()

            '''Fill semua akan di cetak
            DTACetak_Header.FillBy_Primary(TBLCetak_Header, vPrimaryPR_Cancel)
            DTACetak_Item.FillBy_Primary(TBLCetak_Item, vPrimaryPR_Cancel)
            If TBLCetak_Header.DefaultView.Count > 0 Then
                '''looping Header
                TBLCetak_Header.DefaultView.Sort = "tanggal, [no transaksi]"
                'TBLCetak_Header.DefaultView.Sort = "[tanggal] ASC"
                'TBLCetak_Header.DefaultView.Sort = "[no transaksi] ASC"
                For Each DV1 In TBLCetak_Header.DefaultView
                    RowCetak_Header = DV1.Row
                    With RowCetak_Header
                        '''Isi Header
                        rowHeader = TBLHeader.NewTBLHeaderRow
                        rowHeader.Primary = .primary
                        rowHeader.Primary_PR = .primary_pr
                        rowHeader.Document_No = .no_transaksi
                        rowHeader.PR_Number = Me.objPR.GetNoTransaksi(.primary_pr)
                        rowHeader.Tanggal = Format(.tanggal, "dd MMMM yyyy").ToString
                        rowHeader.Reason = .reason
                        rowHeader.Remarks = .remarks
                        TBLHeader.Rows.Add(rowHeader)
                    End With

                    '''Cek dan Looping data Item
                    TBLCetak_Item.DefaultView.RowFilter = Nothing
                    TBLCetak_Item.DefaultView.RowFilter = "[primary transaksi] = " & RowCetak_Header.primary.ToString
                    If TBLCetak_Item.DefaultView.Count > 0 Then
                        For Each DV2 In TBLCetak_Item.DefaultView
                            RowCetak_Item = DV2.Row
                            With RowCetak_Item
                                '''isi item
                                rowItem = TBLItem.NewTblItemRow
                                rowItem.Primary = .primary
                                rowItem.Primary_Transaksi = rowHeader.Primary
                                rowItem.Code_No = ObjProduct.GetProductCode(.primary_code_no)
                                rowItem.Description = ObjProduct.GetDescription(.primary_code_no)

                                Dim DefPrec As Integer = 0
                                DefPrec = ObjProduct.GetAngkaDecimal(.primary_code_no)
                                rowItem.Qty = ObjCD.DigitGrouping(.qty, DefPrec)
                                rowItem.Unit = ObjProduct.GetUnitCode(.primary_code_no)
                                rowItem.Remarks = .remarks
                                TBLItem.Rows.Add(rowItem)
                            End With
                        Next
                    End If
                Next
            End If
            Report.SetDataSource(DS)
            CRV.ReportSource = Report
            CRV.Zoom(95)
            CRV.Refresh()
            CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTACetak_Header.Dispose()
            DTACetak_Item.Dispose()
            TBLCetak_Header.Dispose()
            TBLCetak_Item.Dispose()
            TBLHeader.Dispose()
            TBLItem.Dispose()
        End Try
    End Function

    Public Sub New()
        Core_Module.SetConstring_DTAAntrian(Me.DTAInOUT)
    End Sub
End Class
