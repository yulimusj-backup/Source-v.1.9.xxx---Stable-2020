Option Explicit On
Option Strict Off
Public Class ClsTransaksiPR
    Public vStatus As ClsTransaksiPR.EnumStatus
    Dim objPR_Custom As New CustomSMESS.ClsTransaksiPurchaseRequest

    '''member pada form sendiri
    Private mPrimary As Double = -1
    Private mNoTransaksi As String = ""
    Private mDate As Date = Now
    Private mUrutan As Double
    Private mRFP As Integer = 0
    Private mPrimaryUserOtorisasi As Double = -1
    Private mTimeOtorisasi As Date = #1/1/1950#
    Private mPrimaryRefParty As Double = -1
    Private mRemarks As String = ""
    Private mCreatedBy As String = ""
    Private mCreatedDate As Date = Now
    Private mModifiedBy As String = ""
    Private mModifiedDate As Date = Now
    Private mTimestamp As Date = Now

    'Dim ObjPR_Item As New ClsTransaksiPR_Item
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjRefParty As New Core_Module.ClsMasterRefParty
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct

    '''khusus untuk antrian
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianPurchaseRequest
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian

    'aul
    Dim PickUpNo As String
    Dim PrimaryPickUp As Double

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
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
            Return mDate
        End Get
        Set(ByVal value As Date)
            mDate = value
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

    Public Property RFP() As Integer
        Get
            Return mRFP
        End Get
        Set(ByVal value As Integer)
            mRFP = value
        End Set
    End Property

    Public Property TimeOtorisasi() As Date
        Get
            Return mTimeOtorisasi
        End Get
        Set(ByVal value As Date)
            mTimeOtorisasi = value
        End Set
    End Property

    Public Property PrimaryRefParty() As Double
        Get
            Return mPrimaryRefParty
        End Get
        Set(ByVal value As Double)
            mPrimaryRefParty = value
        End Set
    End Property

    Public Property PrimaryUserOtorisasi() As Double
        Get
            Return mPrimaryUserOtorisasi
        End Get
        Set(ByVal value As Double)
            mPrimaryUserOtorisasi = value
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
    Public Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Public Enum EnumReadyForPurchase
        No
        Yes
    End Enum

    Public Enum EnumStatus
        PR_Pending
        Purchase_On_Progress
        Closed
    End Enum

    Public Enum EnumDeliveryDate
        Uncheklist
        Cheklist
    End Enum

    Public Sub SetDeliveryDateControl(ByVal vDelDate As Date, ByVal vIdentifier As String)
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetDeliveryDateCont(1, vDelDate, vIdentifier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub SetDeliveryDateControl(ByVal vDelDateCont As Boolean, ByVal vDeliveryDate As Date, ByVal vIdentifier As String)
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vDelDateCont = True Then
                DTA.UpdateQuery_SetDeliveryDateCont(1, vDeliveryDate, vIdentifier)
            Else
                DTA.UpdateQuery_SetDeliveryDateCont(0, #1/1/1950#, vIdentifier)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    ''cek di Transaksi sendiri
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

    ''''cek di 2 transaksi 
    'Public Function ApakahSudahAdaDiDuaTransaksi(ByVal vNoTransaksi As String) As Boolean
    '    Try
    '        If Me.GetJmlByNoTransaksi2(vNoTransaksi) > 0 And Me.GetJmlByNoTransaksi3(vNoTransaksi) > 0 Then
    '            Return True
    '        Else
    '            Return False
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    ''''cek di Transaksi PR Cancel
    'Public Function ApakahSudahAda2(ByVal vNoTransaksi As String, ByVal vShowMsg As Boolean) As Boolean
    '    Try
    '        If Me.GetJmlByNoTransaksi2(vNoTransaksi) > 0 Then
    '            If vShowMsg = True Then
    '                MsgBox("The following purchase request data is reference by one or several purchase request cancel, you can’t cancel the authorization/delete this data.", MsgBoxStyle.Information, "Information")
    '            End If
    '            Return True
    '        Else
    '            Return False
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    ''''cek di Transaksi PR pickup order
    'Public Function ApakahSudahAda3(ByVal vNoTransaksi As String, ByVal vShowMsg As Boolean) As Boolean
    '    Try
    '        If Me.GetJmlByNoTransaksi3(vNoTransaksi) > 0 Then
    '            If vShowMsg = True Then
    '                MsgBox("The following purchase request data is reference by one or several purchase request pickup order, you can’t cancel the authorization/delete this data.", MsgBoxStyle.Information, "Information")
    '            End If
    '            Return True
    '        Else
    '            Return False
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    ''''cek yang sudah diotorisasi
    'Public Function ApakahSudahAda4(ByVal vReady As Integer, ByVal vNoTransaksi As String) As Boolean
    '    Try
    '        If Me.GetJmlByReadyForPurchase(vNoTransaksi, vReady) > 0 Then
    '            Return True
    '        Else
    '            Return False
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    ''''cek berapa jml code dalam satu transaksi pr khusus untuk pr cancel or .. 
    'Public Function ApakahSudahAda5(ByVal vPrimaryPR As Double, ByVal vPrimaryCode As Double) As Boolean
    '    Try
    '        If Me.GetJmlByPrimaryCode(vPrimaryPR, vPrimaryCode) > 0 Then
    '            Return True
    '        Else
    '            Return False
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    ''''cek berapa jml ob dalam satu transaksi pr khusus untuk pr cancel or .. 
    'Public Function ApakahSudahAda6(ByVal vPrimaryPR As Double, ByVal vPrimaryCode As Double) As Boolean
    '    Try
    '        If Me.GetOB(vPrimaryPR, vPrimaryCode) > 0 Then
    '            Return True
    '        Else
    '            Return False
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    ''''cek Qty saat Edit
    'Public Function ApakahQtySudahBenar(ByVal vfrm As FrmTransaksiPR, ByVal vPrimaryPR As Double, ByVal QtyCancel As Decimal, ByVal QtyPickUP As Decimal) As Boolean
    '    Dim row As DataGridViewRow
    '    Try
    '        For Each row In vfrm.DgvItem.Rows
    '            Dim OB As Decimal = Me.GetOB(vPrimaryPR, CDbl(row.Cells(vfrm.ColItem_PrimaryProduct.Index).Value))
    '            'Dim QtyCancel As Decimal = Me.GetQtyCancel(vPrimaryPR, CDbl(row.Cells(vfrm.ColItem_PrimaryProduct.Index).Value))
    '            Dim OB_Stlh_Edit As Decimal = 0
    '            Dim MAX As Decimal = 0
    '            Dim Defprec As Decimal = CDec(row.Cells(vfrm.ColItem_PrimaryProduct.Index).Value)
    '            Dim Unit As String = row.Cells(vfrm.ColItem_Unit.Index).Value

    '            OB_Stlh_Edit = OB + QtyCancel + QtyPickUP - CDec(row.Cells(vfrm.ColItem_Qty.Index).Value)
    '            If OB_Stlh_Edit < 0 Then
    '                MAX = OB_Stlh_Edit + CDec(row.Cells(vfrm.ColItem_Qty.Index).Value)
    '                MsgBox("Code No : " & row.Cells(vfrm.ColItem_CodeNo.Index).Value & " " & "maximum value you can edit is" & " " & Me.ObjCD.DigitGrouping(MAX, Defprec) & _
    '                " " & row.Cells(vfrm.ColItem_Unit.Index).Value.ToLower & " ", MsgBoxStyle.Information, "Information")
    '                vfrm.DgvItem.ClearSelection()
    '                vfrm.DgvItem.Rows(row.Index).Selected = True
    '                Return False
    '            End If
    '        Next
    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    'Public Function ApakahDataBenar(ByVal vDGV As DataGridView, ByVal vBaru As Boolean, ByVal vFrm As FrmTransaksiPR) As Boolean

    '    Try
    '        vDGV.AllowUserToAddRows = False
    '        If vBaru = True Then
    '            'jika sudah ada, maka ganti no transaksi
    '            If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
    '                Me.mNoTransaksi = Me.GetNoBaru(Me.mDate)
    '            End If
    '        Else
    '            If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
    '                If Me.mPrimary <> Me.GetPrimary(Me.mNoTransaksi) Then
    '                    Me.mNoTransaksi = Me.GetNoBaru(Me.mDate)
    '                End If
    '            End If
    '        End If

    '        '''Item tidak boleh kosong
    '        If vDGV.Rows.Count = 1 Then
    '            For i As Integer = 0 To vDGV.RowCount - 1
    '                If vDGV.Item(vFrm.ColItem_CodeNo.Index, i).Value = "" Then
    '                    MsgBox("Please input product code", MsgBoxStyle.Information, "Information ")
    '                    Return False
    '                End If
    '            Next
    '        End If

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        vDGV.AllowUserToAddRows = True
    '    End Try
    'End Function

    ''''cek user otorisasi
    'Public Function ApakahUserSama(ByVal vPrimary As Double, ByVal vSalesStaff As String, _
    '                                Optional ByVal vTampilkanPesan As Boolean = True) As Boolean

    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter
    '    Dim UserOtorisasi As String = DTA.ScalarQuery_GetKodeUserOtorisasi(vPrimary)
    '    Try
    '        If UserOtorisasi <> vSalesStaff Then
    '            If vTampilkanPesan = True Then
    '                MsgBox("Only user with ID:" + UserOtorisasi + " can cancel authorize", _
    '                MsgBoxStyle.Information, "Information")
    '            End If
    '            Return False
    '        End If
    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    'Public Function Apakah_UseTracking(ByVal vPrimaryCodeNo As Double, ByVal vUserCode As String, ByVal vPrimaryRefParty As Double, ByVal vRefDate As Date) As Boolean
    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_item2TableAdapter
    '    Dim TBL As New DseTransaksiPR.tbltrans_purchase_req_item2DataTable
    '    Dim row As DseTransaksiPR.tbltrans_purchase_req_item2Row

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA.Fill(TBL, vPrimaryCodeNo, vPrimaryRefParty, vUserCode, vRefDate)
    '        If TBL.Rows.Count = 0 Then
    '            Return True
    '        Else
    '            row = TBL.Rows(0)
    '            If row.track_using = 1 Then
    '                Return True
    '            Else
    '                Return False
    '            End If
    '        End If

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return True
    '    End Try
    'End Function

    ''''cek otorisasi
    Public Function ApakahSudahDiOtorisasi(ByVal vPrimaryPR As Double) As Boolean
        Try
            If Me.GetReadyForPurchase(vPrimaryPR) = ClsTransaksiPR.EnumReadyForPurchase.No Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Public Function HakKhususUser(ByVal vUser As String) As Boolean
    '    Dim ObjUser As New Core_Module.ClsUser

    '    Try
    '        If ObjUser.ApakahMemilikiHakAkses(vUser, "FrmTransaksiPR_Override_Otorisasi") = False Then
    '            Return False
    '        Else
    '            Return True
    '        End If

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjUser = Nothing
    '    End Try
    'End Function

    'Public Function ApakahBolehHapus(ByVal vPrimary As Double, _
    '                                    ByVal vSalesStaff As String, ByVal Frm As FrmTransaksiPR, _
    '                                    Optional ByVal vTampilkanPesan As Boolean = True) As Boolean

    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter
    '    Dim Created_By As String = DTA.ScalarQuery_GetCreatedBy(vPrimary)
    '    Try
    '        If vSalesStaff = "USER" Then Return True
    '        If Created_By <> vSalesStaff Then
    '            If vTampilkanPesan = True Then
    '                MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
    '                MsgBoxStyle.Information, "Information")
    '            End If
    '            Return False
    '        End If
    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    'Public Function ApakahBolehEdit(ByVal frm As FrmTransaksiPR, ByVal vPrimary As Double, ByVal vSalesStaff As String, Optional ByVal vTampilkanPesan As Boolean = True) As Boolean
    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter
    '    Dim Created_By As String = DTA.ScalarQuery_GetCreatedBy(vPrimary)
    '    Try
    '        If vSalesStaff = "USER" Then Return True
    '        If Created_By <> vSalesStaff Then
    '            If vTampilkanPesan = True Then
    '                MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
    '                MsgBoxStyle.Information, "Information")
    '            End If
    '            Return False
    '        End If
    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    'Public Function ApakahBolehHapusItem(ByVal vPrimaryPR As Double, ByVal vPrimaryProduct As Double, _
    '                                        Optional ByVal vShowMsg As Boolean = True) As Boolean
    '    Dim CountPrCancel As Integer = Me.GetJmlItemInPrCancel(vPrimaryPR, vPrimaryProduct)
    '    Dim CountPrPickUp As Integer = Me.GetJmlItemInPickUp(vPrimaryPR, vPrimaryProduct)

    '    Try
    '        If CountPrCancel > 0 Then
    '            If vShowMsg = True Then
    '                MsgBox("The following purchase request data is reference by one or several purchase request cancel, you can’t delete this data.", MsgBoxStyle.Information, "Information")
    '            End If
    '            Return False
    '            Exit Function
    '        End If

    '        If CountPrPickUp > 0 Then
    '            If vShowMsg = True Then
    '                MsgBox("The following purchase request data is reference by one or several purchase request pickup order, you can’t delete this data.", MsgBoxStyle.Information, "Information")
    '            End If
    '            Return False
    '            Exit Function
    '        End If
    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    'Public Function ApakahBolehEditItem(ByVal vPrimaryPR As Double, ByVal vPrimaryProduct As Double, _
    '                                        Optional ByVal vShowMsg As Boolean = True) As Boolean

    '    Dim CountPrCancel As Integer = Me.GetJmlItemInPrCancel(vPrimaryPR, vPrimaryProduct)
    '    Dim CountPrPickUp As Integer = Me.GetJmlItemInPickUp(vPrimaryPR, vPrimaryProduct)

    '    Try
    '        If CountPrCancel > 0 Then
    '            If vShowMsg = True Then
    '                MsgBox("Sorry, can not be changed because it is already used in purchase request cancel", MsgBoxStyle.Information, "Information")
    '            End If
    '            Return False
    '            Exit Function
    '        End If

    '        If CountPrPickUp > 0 Then
    '            If vShowMsg = True Then
    '                MsgBox("Sorry, can not be changed because it is already used in purchase request pickup order", MsgBoxStyle.Information, "Information")
    '            End If
    '            Return False
    '            Exit Function
    '        End If
    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function


    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter
        Dim Primary As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vNoTransaksi).HasValue AndAlso DTA.ScalarQuery_GetPrimary(vNoTransaksi).Value > -1 Then
                Primary = DTA.ScalarQuery_GetPrimary(vNoTransaksi).Value
            End If
            Return Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Public Function GetPrimaryItemByOB(ByVal OB As Double, ByVal vPrimaryPR As Double) As Double
    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
    '    Dim Primary As Double = -1

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If DTA.ScalarQuery_GetPrimaryByOB(OB, vPrimaryPR).HasValue AndAlso DTA.ScalarQuery_GetPrimaryByOB(OB, vPrimaryPR).Value > -1 Then
    '            Primary = DTA.ScalarQuery_GetPrimaryByOB(OB, vPrimaryPR).Value
    '        End If
    '        Return Primary
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    'Public Function GetPrimaryRefParty(ByVal vPrimaryPR As Double) As Double
    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter
    '    Dim Primary As Double = -1

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If DTA.ScalarQuery_GetPrimaryRefParty(vPrimaryPR).HasValue AndAlso DTA.ScalarQuery_GetPrimaryRefParty(vPrimaryPR).Value > -1 Then
    '            Primary = DTA.ScalarQuery_GetPrimaryRefParty(vPrimaryPR).Value
    '        End If
    '        Return Primary
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    'Public Function GetJmlByReadyForPurchase(ByVal vNoTransaksi As String, ByVal vReady As Integer) As Double
    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If String.IsNullOrEmpty(DTA.ScalarQuery_GetJmlReadyForPurchase(vNoTransaksi, vReady)) = False Then
    '            Return DTA.ScalarQuery_GetJmlReadyForPurchase(vNoTransaksi, vReady)
    '        Else
    '            Return 0
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return 0
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Function GetTanggalTransaksi_ByIdentifier(ByVal vIdentifier As String) As Date
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetTanggal_ByIdentifier(vIdentifier)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTanggalTransaksi(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetTanggal(vPrimary)
            If RetVal.HasValue = True Then
                Return RetVal
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Public Function GetTimestamp(ByVal vPrimary As Double) As Date
    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
    '            Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
    '        Else
    '            Return Date.MinValue
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter
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

    Public Function GetRemarks(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetRemarks(vPrimary)
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

    Public Function GetOB(ByVal vPrimaryPR As Double, ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim OB As Double = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetOB(vPrimaryPR, vPrimaryProduct).HasValue AndAlso DTA.ScalarQuery_GetOB(vPrimaryPR, vPrimaryProduct).Value > 0 Then
                OB = DTA.ScalarQuery_GetOB(vPrimaryPR, vPrimaryProduct).Value
            End If
            Return OB
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Public Function GetQtyCancel(ByVal vPrimaryPR As Double, ByVal vPrimaryProduct As Double) As Double
    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_cancel_itemTableAdapter
    '    Dim Qty As Double = 0

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If DTA.ScalarQuery_GetQtyPRCancel(vPrimaryPR, vPrimaryProduct).HasValue AndAlso DTA.ScalarQuery_GetQtyPRCancel(vPrimaryPR, vPrimaryProduct).Value > 0 Then
    '            Qty = DTA.ScalarQuery_GetQtyPRCancel(vPrimaryPR, vPrimaryProduct).Value
    '        End If
    '        Return Qty
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    'Public Function GetQtyPickUp(ByVal vPrimaryPR As Double, ByVal vPrimaryProduct As Double) As Double
    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_pickup_order_itemTableAdapter
    '    Dim Qty As Double = 0

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If DTA.ScalarQuery_GetQtyPRPickUp(vPrimaryPR, vPrimaryProduct).HasValue AndAlso DTA.ScalarQuery_GetQtyPRPickUp(vPrimaryPR, vPrimaryProduct).Value > 0 Then
    '            Qty = DTA.ScalarQuery_GetQtyPRPickUp(vPrimaryPR, vPrimaryProduct).Value
    '        End If
    '        Return Qty
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Function GetJmlByNoTransaksi(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter
        Dim Retval As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJmlNoTransNoLIKE(vNoTransaksi)
            If Retval.HasValue = True Then
                If String.IsNullOrEmpty(Retval) = False Then
                    Return Retval
                Else
                    Return 0
                End If
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Public Function GetJmlByNoTransaksi2(ByVal vNoTransaksi As String) As Double
    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If String.IsNullOrEmpty(DTA.ScalarQuery_GetJmlByNoTransaksi2(vNoTransaksi)) = False Then
    '            Return DTA.ScalarQuery_GetJmlByNoTransaksi2(vNoTransaksi)
    '        Else
    '            Return 0
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return 0
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    'Public Function GetJmlByNoTransaksi3(ByVal vNoTransaksi As String) As Double
    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If String.IsNullOrEmpty(DTA.ScalarQuery_GetJmlByNoTransaksi3(vNoTransaksi)) = False Then
    '            Return DTA.ScalarQuery_GetJmlByNoTransaksi3(vNoTransaksi)
    '        Else
    '            Return 0
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return 0
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Function GetJmlByPrimaryCode(ByVal vPrimaryPR As Double, ByVal vPrimaryCode As Double) As Double
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetJmlByPrimaryCode(vPrimaryPR, vPrimaryCode)) = False Then
                Return DTA.ScalarQuery_GetJmlByPrimaryCode(vPrimaryPR, vPrimaryCode)
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

    'Public Function GetJmlItemInPrCancel(ByVal vPrimaryPR As Double, ByVal vPrimaryProduct As Double) As Double
    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_cancel_itemTableAdapter

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If String.IsNullOrEmpty(DTA.ScalarQuery_GetJmlInPRCancel(vPrimaryProduct, vPrimaryPR)) = False Then
    '            Return DTA.ScalarQuery_GetJmlInPRCancel(vPrimaryProduct, vPrimaryPR)
    '        Else
    '            Return 0
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return 0
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    'Public Function GetJmlItemInPickUp(ByVal vPrimaryPR As Double, ByVal vPrimaryProduct As Double) As Double
    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_pickup_order_itemTableAdapter

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If String.IsNullOrEmpty(DTA.ScalarQuery_GetJmlProductInPickup(vPrimaryPR, vPrimaryProduct)) = False Then
    '            Return DTA.ScalarQuery_GetJmlProductInPickup(vPrimaryPR, vPrimaryProduct)
    '        Else
    '            Return 0
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return 0
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Function GetReadyForPurchase(ByVal vPrimary As Double) As ClsTransaksiPR.EnumReadyForPurchase
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter
        Dim Retval As Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetReadyForPurchase(vPrimary)
            If Retval.HasValue = True Then
                If Retval.Value = 0 Then
                    Return ClsTransaksiPR.EnumReadyForPurchase.No
                ElseIf Retval.Value = 1 Then
                    Return ClsTransaksiPR.EnumReadyForPurchase.Yes
                Else
                    Return ClsTransaksiPR.EnumReadyForPurchase.No
                End If
            End If
            Return ClsTransaksiPR.EnumReadyForPurchase.No
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ClsTransaksiPR.EnumReadyForPurchase.No
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
                                 Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Double
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByNoTransaksi Then
                RetVal = DTA.ScalarQuery_GetJmlByNoTransaksi(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If

            ElseIf vSearchBy = EnumSearchCriteria.ByTanggalTransaksi Then
                RetVal = DTA.ScalarQuery_GetJmlByTangalDariTglSampai(vTanggalDari, vTanggalSampai)
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If

            ElseIf vSearchBy = EnumSearchCriteria.ByRequestBy Then
                RetVal = DTA.ScalarQuery_GetJmlByCode(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If

            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                RetVal = DTA.ScalarQuery_GetJmlByRemarks(vKriteria + "%")
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

    'Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter
    '    Dim TanggalAwal, TanggalAkhir As Date
    '    Dim NoUrut As String
    '    Dim ObjCD As New Core_Module.ClsCekData
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring

    '        TanggalAwal = ObjCD.GetFirstDateofTheMonth(vTanggalTransaksi)
    '        TanggalAkhir = ObjCD.GetLasDateOfTheMonth(vTanggalTransaksi)
    '        If DTA.ScalarQuery_GetMaxUrutan(TanggalAwal, TanggalAkhir).HasValue Then
    '            NoUrut = CStr(DTA.ScalarQuery_GetMaxUrutan(TanggalAwal, TanggalAkhir).Value + 1)
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
    '        Return NoUrut + "/" + "PR" + "/" + CStr(Format(vTanggalTransaksi, "MMM").ToUpper) + "/" + CStr(Format(vTanggalTransaksi, "yyyy"))
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return ""
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function



    'Public Function SimpanBaru1(ByVal vDGVItem As DataGridView, _
    '                                ByVal vUserCode As String, ByVal vFrm As FrmTransaksiPR, _
    '                                ByVal vDT_ItemDelivery As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryDataTable, ByVal vBaru As Boolean) As Boolean
    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If ApakahDataBenar(vDGVItem, vBaru, vFrm) = False Then
    '            Return False
    '        End If

    '        Me.mUrutan = CDbl(Me.mNoTransaksi.Substring(0, 6))

    '        If DTA.Insert(Me.mNoTransaksi, _
    '                      Me.mDate, Me.mUrutan, _
    '                      Me.mRFP, Me.mPrimaryUserOtorisasi, _
    '                      Me.mTimeOtorisasi, Me.mRemarks, Me.mPrimaryRefParty, _
    '                      vUserCode, Now, vUserCode, _
    '                      Now, Now) > 0 Then

    '            Me.mPrimary = Me.GetPrimary(Me.mNoTransaksi)
    '            Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
    '            Me.mCreatedDate = Me.mTimestamp
    '            Me.mModifiedDate = Me.mTimestamp
    '            Me.mCreatedBy = vUserCode
    '            Me.mModifiedBy = vUserCode

    '            'simpan data item
    '            Me.ObjPR_Item.SimpanItem(Me.mPrimary, vFrm.DgvItem, vFrm, vDT_ItemDelivery)

    '            'set order balance
    '            'Me.SetOrderBalance(Me.mPrimary, Me.mRFP)

    '            Return True
    '        Else
    '            Return False
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '        vDT_ItemDelivery.Dispose()
    '    End Try
    'End Function

    'Public Overridable Function SimpanPerubahan1(ByVal vDGVItem As DataGridView, _
    '                                                ByVal vUserCode As String, ByVal vFrm As FrmTransaksiPR, _
    '                                                ByVal vDT_ItemDelivery As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryDataTable, ByVal vBaru As Boolean) As Boolean
    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If ApakahDataBenar(vDGVItem, vBaru, vFrm) = False Then Exit Function

    '        Me.mUrutan = CDbl(Me.mNoTransaksi.Substring(0, 6))
    '        If DTA.Update(Me.mNoTransaksi, _
    '                       Me.mDate, Me.mUrutan, _
    '                      Me.mRFP, Me.mPrimaryUserOtorisasi, _
    '                      Me.mTimeOtorisasi, Me.mRemarks, Me.mPrimaryRefParty, _
    '                      vUserCode, Now, Now, Me.mPrimary) > 0 Then

    '            'update data yang berubah setelah simpan perubahan
    '            Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
    '            Me.mModifiedDate = Me.mTimestamp
    '            Me.mModifiedBy = vUserCode

    '            'hapus data item lama
    '            ObjPR_Item.HapusItem(Me.mPrimary)
    '            'simpan data item baru
    '            Me.ObjPR_Item.SimpanItem(Me.mPrimary, vFrm.DgvItem, vFrm, vDT_ItemDelivery)
    '            'set order balance
    '            'Me.SetOrderBalance(Me.mPrimary, Me.mRFP)
    '            'nilai balikan
    '            Return True
    '        Else
    '            Return False
    '        End If
    '        'End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '        vDT_ItemDelivery.Dispose()
    '    End Try
    'End Function

    'Public Function SimpanBaru(ByVal vDGVItem As DataGridView, _
    '                                ByVal vUserCode As String, ByVal Frm As FrmTransaksiPR, _
    '                                ByVal vDT_ItemDelivery As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryDataTable, ByVal vBaru As Boolean) As Boolean

    '    Dim CheckInterval As Double = 0
    '    Dim TryCounter As Integer = 0
    '    Dim Start As Double = 0
    '    Dim retval As Boolean = True
    '    Dim SW As New Stopwatch

    '    Try
    '        SW.Start()
    '        'ambil no antrian
    '        Me.ObjAntrin.GetNoAntrian(Me.vIdentifier.ToString, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
    '        Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Purchase_Request_SimpanBaru, vUserCode)
    '        'Me.GetNoAntrian()

    '        'coba simpan data baru
    '        Me.Clock_Antri.Reset()
    '        Me.Clock_Antri.Start()
    '        While Start = 0
    '            If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
    '                Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
    '                MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
    '                Exit Function
    '            Else
    '                'hanya mencoba ke database setelah interval waktu tertentu saja
    '                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
    '                    If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
    '                        'clock antri matikan
    '                        Me.Clock_Antri.Stop()
    '                        'sudah saatnya melakukan proses simpan baru
    '                        retval = Me.SimpanBaru1(vDGVItem, vUserCode, Frm, vDT_ItemDelivery, vBaru)
    '                        'hapus data antri untuk transaksi ini
    '                        If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
    '                            MsgBox("Logic error(1), contact your provider when you see this message", _
    '                            MsgBoxStyle.Information, "Information")
    '                        End If
    '                        'nilai balikan
    '                        Return retval
    '                    Else
    '                        'kontrol interval pengecekan
    '                        TryCounter += 1
    '                        CheckInterval = ((TryCounter \ 5) + 1) * 500
    '                        Me.Clock_Antri.Reset()
    '                        Me.Clock_Antri.Start()
    '                        Threading.Thread.Sleep(1000)
    '                    End If
    '                End If
    '            End If
    '        End While
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        SW = Nothing
    '        'vDT_ItemDelivery.Dispose()
    '    End Try
    'End Function

    'Public Overridable Function SimpanPerubahan(ByVal vDGVItem As DataGridView, ByVal vUserCode As String, ByVal frm As FrmTransaksiPR, _
    '                                            ByVal vDT_ItemDelivery As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryDataTable, ByVal vBaru As Boolean) As Boolean
    '    Dim CheckInterval As Double = 0
    '    Dim TryCounter As Integer = 0
    '    Dim Start As Double = 0
    '    Dim retval As Boolean = True
    '    Dim SW As New Stopwatch

    '    Try
    '        SW.Start()

    '        'ambil no antrian
    '        Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, _
    '        Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Purchase_Request_SimpanPerubahan, vUserCode)


    '        'coba simpan data baru
    '        Me.Clock_Antri.Reset()
    '        Me.Clock_Antri.Start()
    '        While Start = 0
    '            'hanya mencoba ke database setelah interval waktu tertentu saja
    '            If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
    '                If SW.Elapsed.TotalSeconds >= Me.ObjAntrin.GetWaitTime Then
    '                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
    '                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
    '                    Exit Function
    '                Else
    '                    If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
    '                        'clock antri matikan
    '                        Me.Clock_Antri.Stop()
    '                        'sudah saatnya melakukan proses simpan baru
    '                        retval = Me.SimpanPerubahan1(vDGVItem, vUserCode, frm, vDT_ItemDelivery, vBaru)
    '                        'hapus data antri untuk transaksi ini
    '                        If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
    '                            MsgBox("Logic error(1), contact your provider when you see this message", _
    '                            MsgBoxStyle.Information, "Information")
    '                        End If
    '                        'nilai balikan
    '                        Return retval
    '                    Else
    '                        'kontrol interval pengecekan
    '                        TryCounter += 1
    '                        CheckInterval = ((TryCounter \ 5) + 1) * 500
    '                        Me.Clock_Antri.Reset()
    '                        Me.Clock_Antri.Start()
    '                        Threading.Thread.Sleep(1000)
    '                    End If
    '                End If
    '            End If
    '        End While
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        SW = Nothing
    '        'vDT_ItemDelivery.Dispose()
    '    End Try
    'End Function

    'Public Overridable Function Hapus1(ByVal vPrimary As Double, ByVal vSalesStaff As String, ByVal vFrm As FrmTransaksiPR) As Boolean
    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        'hapus data header, data item otomatis dihapuskan oleh server
    '        If ApakahBolehHapus(vPrimary, vSalesStaff, vFrm, True) = True Then
    '            If DTA.Delete(vPrimary) > 0 Then
    '                Return True
    '            End If
    '        End If
    '        Return False
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    'Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal usercode As String, ByVal vFrm As FrmTransaksiPR) As Boolean
    '    Dim CheckInterval As Double = 0
    '    Dim TryCounter As Integer = 0
    '    Dim Start As Double = 0
    '    Dim retval As Boolean = True
    '    Dim ObjAntri As New Core_Module.ClsAntrian
    '    Dim SW As New Stopwatch

    '    Try
    '        SW.Start()

    '        'ambil no antrian
    '        ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, _
    '        Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Purchase_Request_Hapus, usercode)

    '        'coba simpan data baru
    '        Me.Clock_Antri.Reset()
    '        Me.Clock_Antri.Start()
    '        While Start = 0
    '            'hanya mencoba ke database setelah interval waktu tertentu saja
    '            If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
    '                If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
    '                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
    '                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
    '                    Exit Function
    '                Else
    '                    If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
    '                        'clock antri matikan
    '                        Me.Clock_Antri.Stop()
    '                        'sudah saatnya melakukan proses simpan baru
    '                        retval = Me.Hapus1(vPrimaryTransaksi, usercode, vFrm)
    '                        'hapus data antri untuk transaksi ini
    '                        If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
    '                            MsgBox("Logic error(1), contact your provider when you see this message", _
    '                            MsgBoxStyle.Information, "Information")
    '                        End If
    '                        'nilai balikan
    '                        Return retval
    '                    Else
    '                        'kontrol interval pengecekan
    '                        TryCounter += 1
    '                        CheckInterval = ((TryCounter \ 5) + 1) * 500
    '                        Me.Clock_Antri.Reset()
    '                        Me.Clock_Antri.Start()
    '                        Threading.Thread.Sleep(1000)
    '                    End If
    '                End If
    '            End If
    '        End While
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        'ObjAntri = Nothing
    '    End Try
    'End Function

    Public Sub SetOrderBalance1(ByVal vPrimaryPR As Double, ByVal vReadyForPurchase As ClsTransaksiPR.EnumReadyForPurchase)
        Dim DTAItem As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim DTACancel As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_cancel_itemTableAdapter
        Dim DTAPickup As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_pickup_order_itemTableAdapter
        Dim DTAPickup1 As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_pickup_order_item1TableAdapter

        Dim TBLItem As New DseTransaksiPR.tbltrans_purchase_req_itemDataTable
        Dim TBLCancel As New DseTransaksiPR.tbltrans_purchase_req_cancel_itemDataTable
        Dim TBlPickup As New DseTransaksiPR.tbltrans_purchase_req_pickup_order_itemDataTable
        Dim TBlPickup1 As New DseTransaksiPR.tbltrans_purchase_req_pickup_order_item1DataTable

        Dim Rowitem As DseTransaksiPR.tbltrans_purchase_req_itemRow
        Dim RowCancel As DseTransaksiPR.tbltrans_purchase_req_cancel_itemRow
        Dim RowPickup As DseTransaksiPR.tbltrans_purchase_req_pickup_order_itemRow
        Dim RowPickup1 As DseTransaksiPR.tbltrans_purchase_req_pickup_order_item1Row

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTACancel.Connection.ConnectionString = DTAItem.Connection.ConnectionString
            DTAPickup.Connection.ConnectionString = DTAItem.Connection.ConnectionString
            DTAPickup1.Connection.ConnectionString = DTAItem.Connection.ConnectionString

            '''Fill setiap tabel 
            DTAItem.FillBy_Primary(TBLItem, vPrimaryPR)
            DTACancel.FillBy_PrimaryPR(TBLCancel, vPrimaryPR)
            DTAPickup.FillBy_PrimaryPR(TBlPickup, vPrimaryPR)
            DTAPickup1.Fill(TBlPickup1, vPrimaryPR)

            If vPrimaryPR = -1 Then
                vReadyForPurchase = EnumReadyForPurchase.No
                vStatus = EnumStatus.PR_Pending
                Exit Sub
            End If

            '''looping tabel PR
            For Each Rowitem In TBLItem
                'set defult value
                Rowitem.OB = Rowitem.qty
                Rowitem.OB_Using = 0
                '''hitung nilai OB + OB Using
                TBLCancel.DefaultView.RowFilter = Nothing
                TBLCancel.DefaultView.RowFilter = "[primary pr] = " & Rowitem.primary_transaksi.ToString & " AND [primary code no] = " & Rowitem.primary_code_no.ToString & ""
                If TBLCancel.DefaultView.Count > 0 Then
                    RowCancel = TBLCancel.DefaultView(0).Row
                    Rowitem.OB -= RowCancel.Qty
                End If

                TBlPickup.DefaultView.RowFilter = Nothing
                TBlPickup.DefaultView.RowFilter = "[primary pr] = " & Rowitem.primary_transaksi.ToString & " AND [primary code no] = " & Rowitem.primary_code_no.ToString & ""
                If TBlPickup.DefaultView.Count > 0 Then
                    RowPickup = TBlPickup.DefaultView(0).Row
                    Rowitem.OB -= RowPickup.Qty
                End If

                'hitung nilai ob using
                If Rowitem.track_using = 1 Then
                    'cari pick up order yang belum di stock out
                    TBlPickup1.DefaultView.RowFilter = "[primary pr] = " & Rowitem.primary_transaksi.ToString & " AND [primary code no] = " & Rowitem.primary_code_no.ToString & ""
                    If TBlPickup1.DefaultView.Count > 0 Then
                        RowPickup1 = TBlPickup1.DefaultView(0).Row
                        Rowitem.OB_Using += RowPickup1.Qty
                    End If
                End If

                If vReadyForPurchase = False Then
                    vStatus = EnumStatus.PR_Pending
                    If Not Rowitem.OB < 0 Then
                        DTAItem.UpdateQuery(Rowitem.qty, 0, vStatus, Rowitem.OB_Using, vPrimaryPR, Rowitem.primary)
                    Else
                        MsgBox("Logic error: Negative order balance found. Please contact your vendor if you see this message.", MsgBoxStyle.Critical, "Error")
                    End If
                Else
                    If Rowitem.OB > 0 Then
                        vStatus = EnumStatus.Purchase_On_Progress
                    Else
                        vStatus = EnumStatus.Closed
                    End If

                    If Rowitem.OB >= 0 Then
                        DTAItem.UpdateQuery(Rowitem.qty, Rowitem.OB, vStatus, Rowitem.OB_Using, vPrimaryPR, Rowitem.primary)
                    End If
                End If
            Next

            '''NOTE : 
            '''hitung ob = pr - cancel - pick up order
            '''simpan ob
            '''set status

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            DTACancel.Dispose()
            DTAPickup.Dispose()
            TBLItem.Dispose()
            TBLCancel.Dispose()
            TBlPickup.Dispose()
            Rowitem = Nothing
            RowCancel = Nothing
            RowPickup = Nothing
        End Try
    End Sub

    'Public Function Authorize(ByVal vTimeAutorize As Date, _
    '                            ByVal vPrimaryAuthorize As Double, _
    '                            ByVal vPrimaryPR As Double, _
    '                            ByVal EnumReadyForPurchase As Integer) As Boolean

    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter

    '    Try
    '        If EnumReadyForPurchase = ClsTransaksiPR.EnumReadyForPurchase.Yes Then
    '            DTA.Connection.ConnectionString = Core_Module.GetConstring

    '            DTA.UpdateQueryAuthorize(EnumReadyForPurchase, vPrimaryAuthorize, vTimeAutorize, vPrimaryPR)
    '            Me.SetOrderBalance(vPrimaryPR, EnumReadyForPurchase)
    '        End If

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    'Public Function CancelAuthorize(ByVal vTimeAutorize As Date, _
    '                            ByVal vPrimaryAuthorize As Double, _
    '                            ByVal vPrimaryPR As Double, _
    '                            ByVal EnumReadyForPurchase As Integer) As Boolean

    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter

    '    Try

    '        If EnumReadyForPurchase = ClsTransaksiPR.EnumReadyForPurchase.No Then
    '            DTA.Connection.ConnectionString = Core_Module.GetConstring

    '            DTA.UpdateQueryAuthorize(EnumReadyForPurchase, vPrimaryAuthorize, vTimeAutorize, vPrimaryPR)
    '            Me.SetOrderBalance(vPrimaryPR, EnumReadyForPurchase)
    '        End If

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function



    'Public Function CariDetail(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
    '                               ByVal vObjCD As Core_Module.ClsCekData, ByVal vFrm As FrmTransaksiPR, _
    '                               ByVal vDT_ItemDelivery As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryDataTable) As Boolean

    '    Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_reqTableAdapter
    '    Dim TBL As New DseTransaksiPR.tbltrans_purchase_reqDataTable
    '    Dim row As DseTransaksiPR.tbltrans_purchase_reqRow

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA.FillBy_Primary(TBL, CDec(vPrimary))
    '        If TBL.Rows.Count > 0 Then
    '            row = TryCast(TBL.Rows(0), DseTransaksiPR.tbltrans_purchase_reqRow)
    '            With row
    '                Me.mPrimary = .primary
    '                Me.mNoTransaksi = .no_transaksi
    '                Me.mDate = .tanggal
    '                Me.mUrutan = .urutan
    '                Me.mRFP = .ready_for_purchase
    '                Me.mPrimaryUserOtorisasi = .primary_user_authorize_by
    '                Me.mTimeOtorisasi = .time_authorize_by
    '                Me.mRemarks = .remarks
    '                Me.mPrimaryRefParty = .primary_ref_party

    '                Me.mCreatedBy = .created_by
    '                Me.mCreatedDate = .created_date
    '                Me.mModifiedBy = .modified_by
    '                Me.mModifiedDate = .modfied_date
    '                Me.mTimestamp = .timestamp
    '            End With

    '            ObjPR_Item.IsiDGVItem(vPrimary, vDGVItem, vFrm, vObjCD, vDT_ItemDelivery)
    '        End If
    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '        TBL.Dispose()
    '        vDT_ItemDelivery.Dispose()
    '    End Try
    'End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
                                        ByVal DGV As DataGridView, ByVal TBL As DseTransaksiPR.tbltrans_purchase_req1DataTable, _
                                        ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
                                        Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
                                        Optional ByVal vTanggalSampai As Date = (#1/1/1950#), _
                                        Optional ByVal vReady As Boolean = False) As Boolean

        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If vReady = False Then
                If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                    QueryResult = DTA.FillBy_GetNoTransaksiLIKE(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                    QueryResult = DTA.FillBy_TglDariTglSampai(TBL, vTanggalDari, vTanggalSampai)
                ElseIf SearchCriteria = EnumSearchCriteria.ByRequestBy Then
                    QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                    QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
                End If
            Else
                If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                    QueryResult = DTA.FillBy_GetNoTransaksiLIKE2(TBL, Criteria + "%", 1)
                ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                    QueryResult = DTA.FillBy_TglDariTglSampai2(TBL, vTanggalDari, vTanggalSampai, 1)
                ElseIf SearchCriteria = EnumSearchCriteria.ByRequestBy Then
                    QueryResult = DTA.FillBy_CodeLIKE2(TBL, 1, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                    QueryResult = DTA.FillBy_RemarksLIKE2(TBL, Criteria + "%", 1)
                End If
            End If

            'update DGV
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(5).Visible = False
                DGV.Columns(1).HeaderText = "Purc. Req No"
                DGV.Columns(2).HeaderText = "Date"
                DGV.Columns(3).HeaderText = "Request By"
                DGV.Columns(4).HeaderText = "Remarks"


                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(1).MinimumWidth = 100
                DGV.Columns(2).MinimumWidth = 130
                DGV.Columns(3).MinimumWidth = 80
                DGV.Columns(4).MinimumWidth = 80
            End If

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 0, MidpointRounding.AwayFromZero) + " miliseconds"
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


    'Public Function CetakPR(ByVal vPrimaryPR As Double, _
    '                        ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer) As Boolean

    '    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    '    Dim ij As New CustomSMESS.FrmPurchaseRiquest_Cetak_Opt
    '    If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Purchase_Request, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
    '        ObjSetting = Nothing
    '        'Return objPR_Custom.CetakPRCustom(vPrimaryPR, CRV)
    '        ij.ShowDialog()
    '        Exit Function
    '    End If

    '    Dim Frm As New FrmTransaksiPR_Cetak
    '    Dim Report As New CRPLaporanPR
    '    Dim DS As New DseTransaksiPR

    '    Dim DTACetak_Header As New DseTransaksiPRTableAdapters.TBLCetak_HeaderTableAdapter
    '    Dim TBLCetak_Header As DseTransaksiPR.TBLCetak_HeaderDataTable = DS.TBLCetak_Header
    '    Dim RowCetak_Header As DseTransaksiPR.TBLCetak_HeaderRow

    '    Dim DTACetak_Item As New DseTransaksiPRTableAdapters.TBLCetak_ItemTableAdapter
    '    Dim TBLCetak_Item As DseTransaksiPR.TBLCetak_ItemDataTable = DS.TBLCetak_Item
    '    Dim RowCetak_Item As DseTransaksiPR.TBLCetak_ItemRow

    '    Dim DTAcetak_Delivery As New DseTransaksiPRTableAdapters.TBLCetak_DeliveryTableAdapter
    '    Dim TBLCetak_Delivery As DseTransaksiPR.TBLCetak_DeliveryDataTable = DS.TBLCetak_Delivery
    '    Dim RowCetak_Delivery As DseTransaksiPR.TBLCetak_DeliveryRow

    '    Dim TBLHeader As DseTransaksiPR.TblHeaderDataTable = DS.TblHeader
    '    Dim TBLItem As DseTransaksiPR.TblItemDataTable = DS.TblItem

    '    Dim rowHeader As DseTransaksiPR.TblHeaderRow
    '    Dim rowItem As DseTransaksiPR.TblItemRow

    '    Dim DV1 As DataRowView
    '    Dim DV2 As DataRowView
    '    'Dim DV3 As DataRowView
    '    Dim filter As String = ""

    '    'Dim FormatDecimal As Integer = 2

    '    Try
    '        DTACetak_Header.Connection.ConnectionString = Core_Module.GetConstring
    '        DTACetak_Item.Connection.ConnectionString = Core_Module.GetConstring
    '        DTAcetak_Delivery.Connection.ConnectionString = Core_Module.GetConstring

    '        Windows.Forms.Cursor.Current = Cursors.WaitCursor

    '        TBLHeader.Rows.Clear()
    '        TBLItem.Rows.Clear()

    '        '''Fill semua akan di cetak
    '        DTACetak_Header.Fill_Primary(TBLCetak_Header, vPrimaryPR)
    '        DTACetak_Item.Fill_Primary(TBLCetak_Item, vPrimaryPR)
    '        DTAcetak_Delivery.Fill_Primary(TBLCetak_Delivery, vPrimaryPR)
    '        '''Jika data header ada
    '        If TBLCetak_Header.DefaultView.Count > 0 Then
    '            '''looping Header
    '            TBLCetak_Header.DefaultView.Sort = "tanggal, [no transaksi]"
    '            For Each DV1 In TBLCetak_Header.DefaultView
    '                RowCetak_Header = DV1.Row
    '                With RowCetak_Header
    '                    '''Isi Header
    '                    rowHeader = TBLHeader.NewTblHeaderRow
    '                    rowHeader.Primary = .primary
    '                    rowHeader.Number_PR = .no_transaksi
    '                    rowHeader.Tanggal = .tanggal

    '                    Dim PrimaryUser As Double = .primary_user_authorize_by
    '                    If PrimaryUser <> -1 Then
    '                        rowHeader.Time_Authorize = " " & Format(.time_authorize_by, "dd MMM yyyy - HH.mm") '& "-" & .time_authorize_by.ToShortTimeString
    '                        rowHeader.User_Authorize = ObjUser.GetKodeUser(.primary_user_authorize_by) & " on" & " "
    '                    Else
    '                        rowHeader.User_Authorize = " - "
    '                        rowHeader.Time_Authorize = ""
    '                    End If
    '                    rowHeader.Request_By = ObjRefParty.GetCode(.primary_ref_party)
    '                    rowHeader.Remarks = .remarks
    '                    TBLHeader.Rows.Add(rowHeader)
    '                End With

    '                '''Cek dan Looping data Item
    '                TBLCetak_Item.DefaultView.RowFilter = Nothing
    '                TBLCetak_Item.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCetak_Header.primary.ToString
    '                If TBLCetak_Item.DefaultView.Count > 0 Then
    '                    For Each DV2 In TBLCetak_Item.DefaultView
    '                        RowCetak_Item = DV2.Row
    '                        With RowCetak_Item
    '                            TBLCetak_Delivery.DefaultView.RowFilter = Nothing
    '                            TBLCetak_Delivery.DefaultView.RowFilter = "[primary transaksi] = " & RowCetak_Header.primary.ToString
    '                            '''isi item
    '                            rowItem = TBLItem.NewTblItemRow
    '                            rowItem.Code = ObjProduct.GetProductCode(.primary_code_no)
    '                            rowItem.Desc = ObjProduct.GetDescription(.primary_code_no)
    '                            rowItem._Decimal = ObjProduct.GetAngkaDecimal(.primary_code_no)
    '                            rowItem.Unit = ObjProduct.GetUnitCode(.primary_unit)
    '                            rowItem.Remarks = .remarks
    '                            rowItem._Decimal = ObjProduct.GetAngkaDecimal(.primary_code_no)
    '                            rowItem.Qty = Me.ObjCD.DigitGrouping(.qty, rowItem._Decimal)
    '                            Dim IntStatus As Integer = RowCetak_Item.status
    '                            If IntStatus = ClsTransaksiPR.EnumStatus.PR_Pending Then
    '                                rowItem.Status = "PR Pending"
    '                            ElseIf IntStatus = ClsTransaksiPR.EnumStatus.Purchase_On_Progress Then
    '                                rowItem.Status = "Purchase On Progress"
    '                            Else
    '                                rowItem.Status = "Closed"
    '                            End If
    '                            TBLItem.Rows.Add(rowItem)
    '                            If .delivery_date_cont = ClsTransaksiPR.EnumDeliveryDate.Cheklist Then
    '                                If .delivery_date = #1/1/1950# Then
    '                                    TBLCetak_Delivery.DefaultView.RowFilter = "[identifier] = '" & RowCetak_Item.identifier.ToString & "'"
    '                                    TBLCetak_Delivery.DefaultView.Sort = "[Delivery Date] ASC"
    '                                    For i As Integer = 0 To TBLCetak_Delivery.DefaultView.Count - 1
    '                                        RowCetak_Delivery = TBLCetak_Delivery.DefaultView(i).Row
    '                                        If i < TBLCetak_Delivery.DefaultView.Count - 1 Then
    '                                            rowItem.Qty2 += Me.ObjCD.DigitGrouping(RowCetak_Delivery.delivery_qty, rowItem._Decimal).ToString & vbNewLine
    '                                            rowItem.Delivery_Date2 += Format(RowCetak_Delivery.delivery_date, "dd MMM yyyy") & vbNewLine
    '                                            rowItem.Delivery_Date = ""
    '                                        Else
    '                                            rowItem.Qty2 += Me.ObjCD.DigitGrouping(RowCetak_Delivery.delivery_qty, rowItem._Decimal).ToString
    '                                            rowItem.Delivery_Date2 += Format(RowCetak_Delivery.delivery_date, "dd MMM yyyy")
    '                                        End If
    '                                    Next
    '                                Else
    '                                    rowItem.Delivery_Date += Format(RowCetak_Item.delivery_date, "dd MMM yyyy")
    '                                End If
    '                            Else
    '                                rowItem.Delivery_Date += "-"
    '                            End If
    '                        End With
    '                    Next
    '                End If
    '            Next
    '        End If

    '        Report.SetDataSource(DS)
    '        CRV.ReportSource = Report
    '        CRV.Zoom(95)
    '        CRV.Refresh()
    '        CRV.Show()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DS.Dispose()
    '        DTACetak_Header.Dispose()
    '        DTACetak_Item.Dispose()
    '        TBLCetak_Header.Dispose()
    '        TBLCetak_Item.Dispose()
    '        TBLHeader.Dispose()
    '        TBLItem.Dispose()
    '        TBLCetak_Delivery.Dispose()
    '        DTAcetak_Delivery.Dispose()
    '    End Try
    'End Function

    Public Function GetRemarksItem(ByVal vPrimaryPR As Double, ByVal vPrimaryCodeNo As Double) As String
        Dim DTAItem As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim RetVal As String = ""

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTAItem.ScalarQuery_getRemarksItem(vPrimaryPR, vPrimaryCodeNo)

            If String.IsNullOrEmpty(RetVal) = True Then
                Return RetVal
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Sub New()
        Core_Module.SetConstring_DTAAntrian(Me.DTAInOUT)
    End Sub

    Public Sub IsikanDeliveryDateData(ByVal vPrimaryPurchReq As Double, ByVal vPrimaryCodeNo As Double, ByVal TBL As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryDataTable)
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryTransaksi_PrimaryCodeNo(TBL, vPrimaryPurchReq, vPrimaryCodeNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function ApakahDeliveryDateControl(ByVal vPrimaryPurchReq As Double, ByVal vPrimaryCodeNo As Double) As Boolean
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim RetVal As Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetDeliveryDateCont(vPrimaryPurchReq, vPrimaryCodeNo)
            If RetVal.HasValue = True Then
                If RetVal.Value = 1 Then
                    Return True
                Else
                    Return False
                End If
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

    Public Function GetOrderBalance_Total(ByVal vPrimaryCodeNo As Double) As Decimal
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetTotalOB(vPrimaryCodeNo)
            If RetVal.HasValue = True Then
                Return RetVal.Value
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

    'aul
    Public Function GetOrderBalance_Total2(ByVal vPrimaryCodeNo As Double) As Decimal
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetTotalOB2(vPrimaryCodeNo)
            If RetVal.HasValue = True Then
                Return RetVal.Value
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

    Public Function GetOrderBalance_Only(ByVal vPrimaryCodeNo As Double) As Decimal
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetOB2(vPrimaryCodeNo)
            If RetVal.HasValue = True Then
                Return RetVal.Value
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

    Public Function GetOrderBalance_Only(ByVal vPrimaryPR As Double, ByVal vPrimaryCodeNo As Double) As Decimal
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetOB(vPrimaryPR, vPrimaryCodeNo)
            If RetVal.HasValue = True Then
                Return RetVal.Value
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

    Public Sub IsiDGVCekStock(ByVal vDGV As DataGridView, ByVal vPrimaryCodeNo As Double, ByVal vFrm As FrmMasterKondisiStock, _
  ByVal vTblPuo2 As DseTransaksiPR.Tbl_PickUpOrder2DataTable)
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_item3TableAdapter
        Dim TBL As New DseTransaksiPR.tbltrans_purchase_req_item3DataTable
        Dim row As DseTransaksiPR.tbltrans_purchase_req_item3Row
        '///

        Dim DTADelivery As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemdelivery1TableAdapter
        Dim TBLDelivery As New DseTransaksiPR.tbltrans_purchase_req_itemdelivery1DataTable
        Dim rowDelivery As DseTransaksiPR.tbltrans_purchase_req_itemdelivery1Row
        Dim PrimaryMax, PrimaryMin As Double
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DummyRemarks As String = ""
        Dim DV As Data.DataRowView

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.Rows.Clear()

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTADelivery.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryCodeNo)

            vFrm.ColPurcReq_PickUpOrder.Visible = False
            If vTblPuo2.Rows.Count > 0 Then
                vTblPuo2.Clear()
            End If

            If TBL.Rows.Count > 0 Then
                PrimaryMax = 0
                PrimaryMin = Double.MaxValue
                For Each row In TBL.Rows
                    If PrimaryMax < row.primary_transaksi Then
                        PrimaryMax = row.primary_transaksi
                    End If
                    If PrimaryMin > row.primary_transaksi Then
                        PrimaryMin = row.primary_transaksi
                    End If
                Next
                DTADelivery.Fill(TBLDelivery, PrimaryMax, PrimaryMin)

                For Each row In TBL.Rows
                    With row
                        DummyRemarks = ""
                        DummyRemarks = .remarks.Trim
                        If DummyRemarks = "" Then
                            If .remarks_item.Trim <> "" Then
                                DummyRemarks = .remarks_item.Trim
                            End If
                        Else
                            If .remarks_item.Trim <> "" Then
                                DummyRemarks = DummyRemarks + " - " + .remarks_item
                            End If
                        End If
                        If DummyRemarks.Length > 250 Then
                            DummyRemarks = DummyRemarks.Substring(0, 250)
                        End If
                    End With

                    Dim _count As Integer = 0

                    If row.delivery_date_cont = 1 Then
                        'urutkan berdasarkan priority level(DESC) dan tgl janji kirim(DESC)
                        'DESC karena asumsi yang dikirim adalah yang High Priority dan Tgl janji kirimnya duluan
                        TBLDelivery.DefaultView.RowFilter = "identifier = '" + row.identifier + "'"
                        TBLDelivery.DefaultView.Sort = "urutan DESC, [delivery date] asc"

                        If TBLDelivery.DefaultView.Count = 0 Then
                            MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            Exit For
                        Else
                            While row.OB > 0 And TBLDelivery.DefaultView.Count > 0
                                rowDelivery = TBLDelivery.DefaultView(0).Row
                                If row.OB >= rowDelivery.delivery_qty Then
                                    With row
                                        Me.TambahBarisDGV(vDGV, vFrm, .Product_Code, .Description, .no_transaksi, .tanggal, rowDelivery.priority_desc, rowDelivery.delivery_qty, _
                                        ._Decimal, ObjCD, .Ref_Party_Desc, DummyRemarks, rowDelivery.delivery_date, 0, "", -1, .primary_transaksi, -1)
                                    End With
                                    row.OB -= rowDelivery.delivery_qty
                                    TBLDelivery.Rows.Remove(rowDelivery)
                                Else
                                    With row
                                        Me.TambahBarisDGV(vDGV, vFrm, .Product_Code, .Description, .no_transaksi, .tanggal, rowDelivery.priority_desc, row.OB, _
                                        ._Decimal, ObjCD, .Ref_Party_Desc, DummyRemarks, rowDelivery.delivery_date, 0, "", -1, .primary_transaksi, -1)
                                    End With
                                    row.OB = 0
                                End If
                            End While
                            While TBLDelivery.DefaultView.Count > 0
                                TBLDelivery.Rows.Remove(TBLDelivery.DefaultView(0).Row)
                            End While
                        End If
                    Else
                        With row
                            Me.TambahBarisDGV(vDGV, vFrm, .Product_Code, .Description, .no_transaksi, .tanggal, "n/a", .OB, ._Decimal, ObjCD, .Ref_Party_Desc, _
                            DummyRemarks, #1/1/1950#, 0, PickUpNo, PrimaryPickUp, .primary_transaksi, -1)
                        End With
                    End If
                Next
            End If

            vDGV.Columns(vFrm.ColPurcReq_Date.Index).DefaultCellStyle.Format = "dd MMM yyyy"
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTADelivery.Dispose()
            TBLDelivery.Dispose()
            ObjCD = Nothing
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    '2018-11-25
    'iwan
    'penyederhanaan fungsi dengan menutup detail order balance using dan detail pick up order
    'di ganti dengan fungsi dengan nama yg sama di atas
    '    Public Sub IsiDGVCekStock(ByVal vDGV As DataGridView, ByVal vPrimaryCodeNo As Double, ByVal vFrm As FrmMasterKondisiStock, _
    '   ByVal vTblPuo2 As DseTransaksiPR.Tbl_PickUpOrder2DataTable)
    '        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_item3TableAdapter
    '        Dim TBL As New DseTransaksiPR.tbltrans_purchase_req_item3DataTable
    '        Dim row As DseTransaksiPR.tbltrans_purchase_req_item3Row

    '        'aul
    '        Dim DtaPUO As New DseTransaksiPRTableAdapters.Tbl_PickUpOrderTableAdapter
    '        Dim TblPUO As New DseTransaksiPR.Tbl_PickUpOrderDataTable
    '        '///

    '        Dim DTADelivery As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemdelivery1TableAdapter
    '        Dim TBLDelivery As New DseTransaksiPR.tbltrans_purchase_req_itemdelivery1DataTable
    '        Dim rowDelivery As DseTransaksiPR.tbltrans_purchase_req_itemdelivery1Row
    '        Dim PrimaryMax, PrimaryMin As Double
    '        Dim RowDGV As DataGridViewRow
    '        Dim ObjCD As New Core_Module.ClsCekData
    '        Dim DummyRemarks As String = ""

    '        Dim rDgv As DataGridViewRow
    '        Dim count, count2 As Integer
    '        Dim Qty, Qty_Dummy, QtyReturn_, OB_, OB_Using As Decimal

    '        Try
    '            Windows.Forms.Cursor.Current = Cursors.WaitCursor
    '            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '            vDGV.Rows.Clear()

    '            DTA.Connection.ConnectionString = Core_Module.GetConstring
    '            DtaPUO.Connection.ConnectionString = Core_Module.GetConstring
    '            DTADelivery.Connection.ConnectionString = Core_Module.GetConstring
    '            DTA.Fill(TBL, vPrimaryCodeNo)

    '            vFrm.ColPurcReq_PickUpOrder.Visible = False
    '            If vTblPuo2.Rows.Count > 0 Then
    '                vTblPuo2.Clear()
    '            End If

    '            If TBL.Rows.Count > 0 Then
    '                PrimaryMax = 0
    '                PrimaryMin = Double.MaxValue
    '                For Each row In TBL.Rows
    '                    If PrimaryMax < row.primary_transaksi Then
    '                        PrimaryMax = row.primary_transaksi
    '                    End If
    '                    If PrimaryMin > row.primary_transaksi Then
    '                        PrimaryMin = row.primary_transaksi
    '                    End If
    '                Next
    '                DTADelivery.Fill(TBLDelivery, PrimaryMax, PrimaryMin)
    '                DtaPUO.FillBy_PrimaryCodeNo(TblPUO, vPrimaryCodeNo)

    '                For Each row In TBL.Rows
    '                    With row
    '                        DummyRemarks = ""
    '                        DummyRemarks = .remarks.Trim
    '                        If DummyRemarks = "" Then
    '                            If .remarks_item.Trim <> "" Then
    '                                DummyRemarks = .remarks_item.Trim
    '                            End If
    '                        Else
    '                            If .remarks_item.Trim <> "" Then
    '                                DummyRemarks = DummyRemarks + " - " + .remarks_item
    '                            End If
    '                        End If
    '                        If DummyRemarks.Length > 250 Then
    '                            DummyRemarks = DummyRemarks.Substring(0, 250)
    '                        End If
    '                    End With

    '                    Dim _count As Integer = 0

    '                    If row.delivery_date_cont = 1 Then
    '                        'kosongkan dulu pick up order (default)
    '                        PickUpNo = "-"
    '                        PrimaryPickUp = -2

    '                        'urutkan berdasarkan priority level(DESC) dan tgl janji kirim(DESC)
    '                        'DESC karena asumsi yang dikirim adalah yang High Priority dan Tgl janji kirimnya duluan
    '                        TBLDelivery.DefaultView.RowFilter = "identifier = '" + row.identifier + "'"
    '                        TBLDelivery.DefaultView.Sort = "urutan DESC, [delivery date] asc"

    '                        If TBLDelivery.DefaultView.Count = 0 Then
    '                            MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
    '                            Exit For
    '                        Else
    '                            Qty = 0
    '                            OB_Using = 0
    '                            Qty_Dummy = 0
    '                            If row.track_using = 1 Then
    '                                If row.qty <> row.OB + row.OB_Using Then
    '                                    row.OB_Using = row.qty - row.OB
    '                                    row.Total_OB = row.qty
    '                                End If
    '                            End If
    '                            While TBLDelivery.DefaultView.Count > 0 And row.Total_OB > 0
    '                                _count += 1
    '                                rowDelivery = TBLDelivery.DefaultView(0).Row
    '                                If row.track_using = 0 Then
    '                                    GoTo fungsi
    '                                End If
    '                                If rowDelivery.delivery_qty <= row.Total_OB Then
    '                                    'tambahkan baris sejumlah delivery qty
    '                                    If rowDelivery.delivery_qty <= row.OB_Using Then
    '                                        If rowDelivery.delivery_qty < row.OB_Using Then
    '                                            OB_Using = rowDelivery.delivery_qty
    '                                        Else
    '                                            OB_Using = row.OB_Using
    '                                        End If
    '                                        If OB_Using > 0 Then
    '                                            QtyReturn_ = Me.GetPickUpOrderNo(TblPUO, rowDelivery, row.primary_transaksi, rowDelivery.delivery_qty, vTblPuo2, _count)
    '                                            OB_Using = OB_Using - QtyReturn_
    '                                        End If
    '                                        If OB_Using <= 0 Then
    '                                            GoTo skipTambahBaris1
    '                                        End If
    '                                        Me.TambahBarisDGV(vDGV, vFrm, row.Product_Code, row.Description, row.no_transaksi, row.tanggal, rowDelivery.priority_desc, _
    '                                        0, row._Decimal, ObjCD, row.Ref_Party_Desc, DummyRemarks, rowDelivery.delivery_date, OB_Using, PickUpNo, _
    '                                        PrimaryPickUp, row.primary_transaksi, rowDelivery.primary)
    'skipTambahBaris1:
    '                                    Else
    'fungsi:
    '                                        If row.track_using = 1 Then
    '                                            Qty = rowDelivery.delivery_qty - row.OB_Using
    '                                        Else
    '                                            count2 += 1
    '                                            If count2 = 1 Then
    '                                                Qty_Dummy = rowDelivery.delivery_qty - (row.qty - row.OB)
    '                                                Qty = rowDelivery.delivery_qty - (row.qty - row.OB)
    '                                                If Qty < 0 Then
    '                                                    Qty = 0
    '                                                End If
    '                                            Else
    '                                                If Qty_Dummy < 0 Then
    '                                                    Qty = rowDelivery.delivery_qty + Qty_Dummy
    '                                                    Qty_Dummy = rowDelivery.delivery_qty + Qty_Dummy
    '                                                    If Qty < 0 Then
    '                                                        Qty = 0
    '                                                    End If
    '                                                Else
    '                                                    Qty = rowDelivery.delivery_qty
    '                                                End If
    '                                            End If

    '                                            If Qty <= 0 Then
    '                                                GoTo skipTambahBaris2
    '                                            End If

    '                                        End If
    '                                        If row.OB_Using > 0 Then
    '                                            QtyReturn_ = Me.GetPickUpOrderNo(TblPUO, rowDelivery, row.primary_transaksi, rowDelivery.delivery_qty, vTblPuo2, _count)
    '                                            OB_Using = row.OB_Using - QtyReturn_
    '                                        Else
    '                                            PickUpNo = "-"
    '                                            PrimaryPickUp = -2
    '                                            OB_Using = row.OB_Using
    '                                        End If
    '                                        Me.TambahBarisDGV(vDGV, vFrm, row.Product_Code, row.Description, row.no_transaksi, row.tanggal, rowDelivery.priority_desc, _
    '                                        Qty, row._Decimal, ObjCD, row.Ref_Party_Desc, DummyRemarks, rowDelivery.delivery_date, OB_Using, _
    '                                        PickUpNo, PrimaryPickUp, row.primary_transaksi, rowDelivery.primary)
    'skiptambahbaris2:
    '                                    End If

    '                                    'set ulang parameter looping
    '                                    If row.track_using = 1 Then
    '                                        row.Total_OB -= rowDelivery.delivery_qty
    '                                        row.OB_Using -= rowDelivery.delivery_qty
    '                                        If row.OB_Using < 0 Then
    '                                            row.OB = row.OB + row.OB_Using
    '                                            row.OB_Using = 0
    '                                        End If
    '                                    Else
    '                                        row.Total_OB -= Qty
    '                                    End If

    '                                    TBLDelivery.Rows.Remove(rowDelivery)
    '                                Else
    '                                    'tambahkan baris sejumlah total ob
    '                                    If row.OB_Using > 0 Then
    '                                        QtyReturn_ = Me.GetPickUpOrderNo(TblPUO, rowDelivery, row.primary_transaksi, rowDelivery.delivery_qty, vTblPuo2, _count)
    '                                        OB_Using = row.OB_Using - QtyReturn_
    '                                    Else
    '                                        PickUpNo = "-"
    '                                        PrimaryPickUp = -2
    '                                        OB_Using = row.OB_Using
    '                                    End If
    '                                    Me.TambahBarisDGV(vDGV, vFrm, row.Product_Code, row.Description, row.no_transaksi, row.tanggal, rowDelivery.priority_desc, _
    '                                    row.OB, row._Decimal, ObjCD, row.Ref_Party_Desc, DummyRemarks, rowDelivery.delivery_date, OB_Using, PickUpNo, _
    '                                    PrimaryPickUp, row.primary_transaksi, rowDelivery.primary)
    '                                    rowDelivery.delivery_qty -= row.Total_OB
    '                                    row.Total_OB = 0
    '                                    row.OB = 0
    '                                    row.OB_Using = 0
    '                                End If
    '                            End While
    '                        End If
    '                    Else
    '                        _count += 1
    '                        OB_Using = 0
    '                        If row.track_using = 1 Then
    '                            If row.qty <> row.OB + row.OB_Using Then
    '                                row.OB_Using = row.qty - row.OB - row.qty_cancel
    '                            End If
    '                        End If

    '                        'langsung tambahkan ke dgv
    '                        If row.OB_Using > 0 Then
    '                            QtyReturn_ = Me.GetPickUpOrderNo(TblPUO, rowDelivery, row.primary_transaksi, row.OB, vTblPuo2, _count, row.qty)
    '                            OB_Using = row.OB_Using - QtyReturn_
    '                        Else
    '                            PickUpNo = "-"
    '                            PrimaryPickUp = -2
    '                            OB_Using = row.OB_Using
    '                        End If
    '                        With row
    '                            Me.TambahBarisDGV(vDGV, vFrm, .Product_Code, .Description, .no_transaksi, .tanggal, "n/a", .OB, ._Decimal, ObjCD, .Ref_Party_Desc, _
    '                            DummyRemarks, #1/1/1950#, OB_Using, PickUpNo, PrimaryPickUp, .primary_transaksi, -1)
    '                        End With
    '                    End If
    '                Next

    '                count = 0
    '                For i As Integer = 0 To vDGV.RowCount - 1
    '                    rDgv = vDGV.Rows(i)
    '                    If rDgv.Cells(vFrm.ColPurcReq_PrimaryPickUpOrder.Index).Value <> -2 Then
    '                        count += 1
    '                        If count = 1 Then
    '                            vFrm.ColPurcReq_PickUpOrder.Visible = True
    '                        End If
    '                    End If
    '                Next
    '            End If

    '            vDGV.Columns(vFrm.ColPurcReq_Date.Index).DefaultCellStyle.Format = "dd MMM yyyy"
    '            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '        Catch ex As Exception
    '            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Finally
    '            DTA.Dispose()
    '            TBL.Dispose()
    '            DtaPUO.Dispose()
    '            TblPUO.Dispose()
    '            DTADelivery.Dispose()
    '            TBLDelivery.Dispose()
    '            ObjCD = Nothing
    '            Windows.Forms.Cursor.Current = Cursors.Arrow
    '        End Try
    '    End Sub

    Private Function GetPickUpOrderNo(ByVal vTblPUO As DseTransaksiPR.Tbl_PickUpOrderDataTable, ByVal RowDelivery As DseTransaksiPR.tbltrans_purchase_req_itemdelivery1Row, _
    ByVal vPrimaryPR As Double, ByVal vDelDate_Qty As Decimal, ByVal vTblPuo2 As DseTransaksiPR.Tbl_PickUpOrder2DataTable, ByVal vCount As Integer, Optional ByVal vQty As Decimal = 0) As Decimal
        Dim RowPUO As DseTransaksiPR.Tbl_PickUpOrderRow
        Dim RowPUO2 As DseTransaksiPR.Tbl_PickUpOrder2Row
        Dim count As Integer = 0
        Dim QtyDel As Decimal = 0
        Dim QtyPickUp As Decimal = 0
        Dim QtyReturn As Decimal = 0
        Dim QtyTemp As Decimal = 0

        Try
            PickUpNo = "-"
            PrimaryPickUp = -2

            vTblPUO.DefaultView.RowFilter = Nothing
            vTblPUO.DefaultView.RowFilter = "[Primary PR] = " & vPrimaryPR
            If vTblPUO.DefaultView.Count > 0 Then
                count = 0
                If vQty > 0 Then
                    QtyDel = vQty
                Else
                    QtyDel = RowDelivery.delivery_qty
                End If
                For i As Integer = 0 To vTblPUO.DefaultView.Count - 1
                    RowPUO = vTblPUO.DefaultView(0).Row

                    vTblPuo2.DefaultView.RowFilter = Nothing
                    If vQty > 0 Then
                        vTblPuo2.DefaultView.RowFilter = "[Primary PR] = " & vPrimaryPR & " AND [Primary Pick Up] = " & RowPUO.Primary_Pick_Up_Order & " AND [Primary PR Item Delivery] = " & -1
                    Else
                        vTblPuo2.DefaultView.RowFilter = "[Primary PR] = " & vPrimaryPR & " AND [Primary Pick Up] = " & RowPUO.Primary_Pick_Up_Order & " AND [Primary PR Item Delivery] = " & RowDelivery.primary
                    End If

                    If RowPUO.primary_stock_out_adjust = -1 Then
                        If vTblPuo2.DefaultView.Count < 1 Then
                            RowPUO2 = vTblPuo2.NewTbl_PickUpOrder2Row
                            RowPUO2.Pick_Up_No = RowPUO.No_Pick_Up_Order
                            RowPUO2.Primary_Pick_Up = RowPUO.Primary_Pick_Up_Order
                            RowPUO2.Primary_PR = vPrimaryPR
                            If vQty > 0 Then
                                RowPUO2.Primary_PR_Item_Delivery = -1
                            Else
                                RowPUO2.Primary_PR_Item_Delivery = RowDelivery.primary
                            End If
                            vTblPuo2.Rows.Add(RowPUO2)
                        End If

                        count += 1
                    End If

                    QtyPickUp = RowPUO.Qty_Pick_Up_Order
                    RowPUO.Qty_Pick_Up_Order -= QtyDel
                    QtyDel -= QtyPickUp

                    If i < vCount Then
                        If RowPUO.primary_stock_out_adjust <> -1 Then
                            If count > 0 Then
                                If RowPUO.Qty_Pick_Up_Order <= 0 Then
                                    QtyReturn += QtyPickUp
                                Else
                                    QtyReturn += RowPUO.Qty_Pick_Up_Order
                                End If
                            Else
                                QtyReturn += QtyPickUp
                            End If
                        End If
                    Else
                        If RowPUO.primary_stock_out_adjust <> -1 Then
                            If QtyTemp < 0 Then
                                QtyReturn = QtyReturn - QtyTemp
                                QtyTemp = 0
                            End If
                        End If
                    End If

                    If count > 1 Then
                        PickUpNo = "Multiple"
                        PrimaryPickUp = -1
                    Else
                        If RowPUO.primary_stock_out_adjust = -1 Then
                            PickUpNo = RowPUO.No_Pick_Up_Order
                            PrimaryPickUp = RowPUO.Primary_Pick_Up_Order
                        End If
                    End If

                    If RowPUO.Qty_Pick_Up_Order <= 0 Then
                        QtyTemp = RowPUO.Qty_Pick_Up_Order
                        vTblPUO.Rows.Remove(RowPUO)
                    End If

                    If QtyDel <= 0 Then
                        Return QtyReturn
                    End If
                Next
            Else
                PickUpNo = "-"
                PrimaryPickUp = -2
            End If

            Return QtyReturn
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Sub IsiDGVPickUpList(ByVal vDgv As DataGridView, ByVal vPrimaryPR As Double, _
    ByVal vPrimaryPrItemDel As String, ByVal vFrm As FrmMasterKondisiStock_PickUpOrderList, _
    ByVal vTblPuo2_ As DseTransaksiPR.Tbl_PickUpOrder2DataTable)
        Dim vRow As DseTransaksiPR.Tbl_PickUpOrder2Row
        Dim rowDGV As DataGridViewRow

        Try
            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            vTblPuo2_.DefaultView.RowFilter = Nothing
            vTblPuo2_.DefaultView.RowFilter = "[Primary PR] = " & vPrimaryPR & "AND [Primary PR Item Delivery] = " & vPrimaryPrItemDel
            If vTblPuo2_.DefaultView.Count > 0 Then
                For i As Integer = 0 To vTblPuo2_.DefaultView.Count - 1
                    vRow = vTblPuo2_.DefaultView(i).Row
                    vDgv.Rows.Add()
                    rowDGV = vDgv.Rows(vDgv.Rows.Count - 1)
                    rowDGV.Cells(vFrm.Col_PrimaryPickUpOrder.Index).Value = vRow.Primary_Pick_Up
                    rowDGV.Cells(vFrm.Col_PickUpOrderNo.Index).Value = vRow.Pick_Up_No
                Next
            End If

            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahBarisDGV(ByVal vDGV As DataGridView, ByVal vFRM As FrmMasterKondisiStock, ByVal vCodeNo As String, ByVal vDesc As String, _
    ByVal vPRNo As String, ByVal vPRDate As Date, ByVal vPriorityLevel As String, ByVal vQtyOB As Decimal, ByVal vDefPrec As Integer, _
    ByVal vObjCD As Core_Module.ClsCekData, ByVal vRefParty As String, ByVal vRemarks As String, ByVal vPromiseToDeliver As Date, ByVal vQtyOBUsing As Decimal, _
    ByVal vPickUpNo As String, ByVal vPrimaryPickUp As Double, ByVal vPrimaryPR As Double, ByVal vPrimaryPRItemDel As Double)
        Dim rowDGV As DataGridViewRow

        Try
            vDGV.Rows.Add()
            rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
            With rowDGV
                .Cells(vFRM.ColPurcReq_CodeNo.Index).Value = vCodeNo
                .Cells(vFRM.ColPurcReq_Description.Index).Value = vDesc
                .Cells(vFRM.ColPurcReq_Date.Index).Value = vPRDate
                .Cells(vFRM.ColPurcReq_DocumentNo.Index).Value = vPRNo
                .Cells(vFRM.ColPurcReq_PriorityLevel.Index).Value = vPriorityLevel
                If vQtyOB > 0 Then
                    .Cells(vFRM.ColPurcReq_QtyOB.Index).Value = vObjCD.DigitGrouping(vQtyOB, vDefPrec)
                Else
                    .Cells(vFRM.ColPurcReq_QtyOB.Index).Value = "-"
                End If
                .Cells(vFRM.ColPurcReq_RefParty.Index).Value = vRefParty
                .Cells(vFRM.ColPurcReq_Remarks.Index).Value = vRemarks
                If vPromiseToDeliver = #1/1/1950# Then
                    .Cells(vFRM.ColPurcReq_PromiseToDeliver.Index).Value = "n/a"
                Else
                    .Cells(vFRM.ColPurcReq_PromiseToDeliver.Index).Value = Format(vPromiseToDeliver, "dd MMM yyyy")
                End If
                If vQtyOBUsing > 0 Then
                    .Cells(vFRM.ColPurcReq_BalanceQtyUsing.Index).Value = vObjCD.DigitGrouping(vQtyOBUsing, vDefPrec)
                Else
                    .Cells(vFRM.ColPurcReq_BalanceQtyUsing.Index).Value = "-"
                End If
                'aul
                .Cells(vFRM.ColPurcReq_PickUpOrder.Index).Value = vPickUpNo
                .Cells(vFRM.ColPurcReq_PrimaryPickUpOrder.Index).Value = vPrimaryPickUp
                .Cells(vFRM.ColPurcReq_PrimaryPR.Index).Value = vPrimaryPR
                .Cells(vFRM.ColPurcReq_PrimaryPrItemDel.Index).Value = vPrimaryPRItemDel
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetJmlByIdentifier(ByVal vIdentifier As String) As Decimal
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlByIdentifier(vIdentifier)
            If RetVal.HasValue = True Then
                Return RetVal.Value
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

    Public Sub UpdateQtySpare(ByVal vPrimaryPR As Double, ByVal vPrimaryProduct As Double, ByVal vQtySpare As Decimal)
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SpareQty(vQtySpare, vPrimaryPR, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetTglDeliveryMinimum(ByVal vPrimaryPR As Double) As Date
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemdelivery2TableAdapter
        Dim TBL As New DseTransaksiPR.tbltrans_purchase_req_itemdelivery2DataTable
        Dim row As DseTransaksiPR.tbltrans_purchase_req_itemdelivery2Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryPR)

            If TBL.Rows.Count = 0 Then
                Return #1/1/1950#
            Else
                row = TBL.Rows(0)
                Return row.delivery_date
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        End Try
    End Function

    Public Sub SimpanItemDetailDelivery(ByVal vTBLDelivery As DseTransaksiMaterialPlanning2.tbltrans_purchase_req_itemdeliveryDataTable)
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
        Dim RowDelivery As DseTransaksiMaterialPlanning2.tbltrans_purchase_req_itemdeliveryRow
        Dim vPrimaryPR As Double = -1
        Dim Dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each RowDelivery In vTBLDelivery.Rows
                With RowDelivery
                    Dummy = RowDelivery.identifier
                    vPrimaryPR = Dummy.Substring(0, Dummy.IndexOf("-"))
                    DTA.Insert(vPrimaryPR, .identifier, .delivery_qty, .delivery_date, .remarks, .delivery_priority)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HapusItemDetailDelivery(ByVal vIdentifierPrItem As String)
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery1(vIdentifierPrItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetJmlOrder(ByVal vIdentifierPrItem As String) As Decimal
        Dim DTA As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlByIdentifier(vIdentifierPrItem)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
