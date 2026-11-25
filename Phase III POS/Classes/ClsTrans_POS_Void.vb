Option Explicit On
Option Strict Off
Imports Core_Module

Public Class ClsTrans_POS_Void

    Public Function GetConstring() As String
        Try
            Return "server=" + My.Settings.server + ";user id=" + My.Settings.user_id + ";Password=" + My.Settings.password + _
            ";persist security info=True;database=" + My.Settings.database_name + ";port=" + My.Settings.port_number + ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=localhost;user id=root;Password=123;persist security info=True;database=core_db;port=3307"
        End Try
    End Function

    Public Function HapusByPrimaryPOS(ByVal vPrimaryPOS As Double) As Boolean
        Dim DTA As New DseTrans_POS_VoidTableAdapters.tbltrans_session_cabang_itemvoid1TableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            DTA.DeleteQuery_PrimaryTransPOS(vPrimaryPOS)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetReason(ByVal vPrimaryPOS As Double, ByVal vPrimaryProduct As Double, ByVal vPrimarySession As Double, _
    Optional ByVal vPrimaryUser As Double = -1) As String
        Dim DTA As New DseTrans_POS_VoidTableAdapters.tbltrans_session_cabang_itemvoidTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetReason(vPrimaryPOS, vPrimaryProduct, vPrimarySession, vPrimaryUser)) = False Then
                Return DTA.ScalarQuery_GetReason(vPrimaryPOS, vPrimaryProduct, vPrimarySession, vPrimaryUser)
            Else
                Return ""
            End If
        Catch ex As Exception
            Return ""
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetQty(ByVal vPrimaryPOS As Double, ByVal vPrimaryProduct As Double, ByVal vPrimarySession As Double, _
    Optional ByVal vPrimaryUser As Double = -1) As Decimal
        Dim DTA As New DseTrans_POS_VoidTableAdapters.tbltrans_session_cabang_itemvoidTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetQty(vPrimarySession, vPrimaryProduct, vPrimaryPOS, vPrimaryUser).HasValue Then
                Return DTA.ScalarQuery_GetQty(vPrimarySession, vPrimaryProduct, vPrimaryPOS, vPrimaryUser).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanItem(ByVal vPrimaryPOS As Double, ByVal vPrimarySession As Double, ByVal vDGVItemProduct As DataGridView, _
     ByVal i As FrmTrans_Session_OpenClose, ByVal vReason As String, ByVal vPrimaryUser As Double) As Boolean
        Dim DTAItemP As New DseTrans_POS_VoidTableAdapters.tbltrans_session_cabang_itemvoidTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData

        Try
            DTAItemP.Connection.ConnectionString = Me.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            For Each rowDGV In vDGVItemProduct.Rows
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColPrimaryDisc1.Index).Value) = True Then
                    rowDGV.Cells(i.ColPrimaryDisc1.Index).Value = "-1"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColPrimaryDisc2.Index).Value) = True Then
                    rowDGV.Cells(i.ColPrimaryDisc2.Index).Value = "-1"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColPrimaryDisc3.Index).Value) = True Then
                    rowDGV.Cells(i.ColPrimaryDisc3.Index).Value = "-1"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscAmount1.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscAmount1.Index).Value = "0"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscAmount2.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscAmount2.Index).Value = "0"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscAmount3.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscAmount3.Index).Value = "0"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscString1.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscString1.Index).Value = ""
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscString2.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscString2.Index).Value = ""
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscString3.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscString3.Index).Value = ""
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColExpiry.Index).Value) = True Then
                    rowDGV.Cells(i.ColExpiry.Index).Value = "1950/01/01"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColPrimaryLocation.Index).Value) = True Then
                    rowDGV.Cells(i.ColPrimaryLocation.Index).Value = "-1"
                End If
                With rowDGV
                    DTAItemP.Insert(vPrimarySession, .Cells(i.ColPrimaryCodeNo.Index).Value, vPrimaryPOS, .Cells(i.ColQty.Index).Value, .Cells(i.ColUnitPrice.Index).Value, _
                    .Cells(i.ColDiscString.Index).Value, .Cells(i.ColDisc_Amount.Index).Value, .Cells(i.ColNettPrice.Index).Value, .Cells(i.ColSubTotal.Index).Value, _
                    .Cells(i.ColPrimaryDisc1.Index).Value, .Cells(i.ColDiscString1.Index).Value, .Cells(i.ColDiscAmount1.Index).Value, _
                    .Cells(i.ColPrimaryDisc2.Index).Value, .Cells(i.ColDiscString2.Index).Value, .Cells(i.ColDiscAmount2.Index).Value, _
                    .Cells(i.ColPrimaryDisc3.Index).Value, .Cells(i.ColDiscString3.Index).Value, .Cells(i.ColDiscAmount3.Index).Value, _
                    .Cells(i.ColPrimaryLocation.Index).Value, .Cells(i.ColBatchNumber.Index).Value, CDate(.Cells(i.ColExpiry.Index).Value), _
                    .Cells(i.ColPrimaryUnit.Index).Value, vPrimaryUser, vReason)
                End With
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemP.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function SimpanItem(ByVal vPrimaryPOS As Double, ByVal vPrimarySession As Double, ByVal vDGVItemProduct As DataGridView, _
      ByVal i As FrmTrans_POS, ByVal vReason As String, ByVal vPrimaryUser As Double) As Boolean
        Dim DTAItemP As New DseTrans_POS_VoidTableAdapters.tbltrans_session_cabang_itemvoidTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData

        Try
            DTAItemP.Connection.ConnectionString = Me.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            For Each rowDGV In vDGVItemProduct.Rows
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColPrimaryDisc1.Index).Value) = True Then
                    rowDGV.Cells(i.ColPrimaryDisc1.Index).Value = "-1"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColPrimaryDisc2.Index).Value) = True Then
                    rowDGV.Cells(i.ColPrimaryDisc2.Index).Value = "-1"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColPrimaryDisc3.Index).Value) = True Then
                    rowDGV.Cells(i.ColPrimaryDisc3.Index).Value = "-1"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscAmount1.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscAmount1.Index).Value = "0"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscAmount2.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscAmount2.Index).Value = "0"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscAmount3.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscAmount3.Index).Value = "0"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscString1.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscString1.Index).Value = ""
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscString2.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscString2.Index).Value = ""
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColDiscString3.Index).Value) = True Then
                    rowDGV.Cells(i.ColDiscString3.Index).Value = ""
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColExpiry.Index).Value) = True Then
                    rowDGV.Cells(i.ColExpiry.Index).Value = "1950/01/01"
                End If
                If String.IsNullOrEmpty(rowDGV.Cells(i.ColPrimaryLocation.Index).Value) = True Then
                    rowDGV.Cells(i.ColPrimaryLocation.Index).Value = "-1"
                End If
                With rowDGV
                    DTAItemP.Insert(vPrimarySession, .Cells(i.ColPrimaryCodeNo.Index).Value, vPrimaryPOS, .Cells(i.ColQty.Index).Value, .Cells(i.ColUnitPrice.Index).Value, _
                    .Cells(i.ColDiscString.Index).Value, .Cells(i.ColDisc_Amount.Index).Value, .Cells(i.ColNettPrice.Index).Value, .Cells(i.ColSubTotal.Index).Value, _
                    .Cells(i.ColPrimaryDisc1.Index).Value, .Cells(i.ColDiscString1.Index).Value, .Cells(i.ColDiscAmount1.Index).Value, _
                    .Cells(i.ColPrimaryDisc2.Index).Value, .Cells(i.ColDiscString2.Index).Value, .Cells(i.ColDiscAmount2.Index).Value, _
                    .Cells(i.ColPrimaryDisc3.Index).Value, .Cells(i.ColDiscString3.Index).Value, .Cells(i.ColDiscAmount3.Index).Value, _
                    .Cells(i.ColPrimaryLocation.Index).Value, .Cells(i.ColBatchNumber.Index).Value, CDate(.Cells(i.ColExpiry.Index).Value), _
                    .Cells(i.ColPrimaryUnit.Index).Value, vPrimaryUser, vReason)
                End With
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItemP.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal vPrimaryUser As Double, ByVal vPrimaryPOS As Double, _
    ByVal vPrimarySession As Double, ByVal vPrimaryProduct As Double) As Double
        Dim DTAItem As New DseTrans_POS_VoidTableAdapters.tbltrans_session_cabang_itemvoid1TableAdapter
        Try
            DTAItem.Connection.ConnectionString = Me.GetConstring

            If DTAItem.ScalarQuery_PrimaryAll(vPrimarySession, vPrimaryProduct, vPrimaryPOS, vPrimaryUser).HasValue Then
                Return DTAItem.ScalarQuery_PrimaryAll(vPrimarySession, vPrimaryProduct, vPrimaryPOS, vPrimaryUser).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function SimpanItemViodSatuSatu(ByVal vQty As Decimal, _
    ByVal vUnitPrice As Decimal, ByVal vDiscString As String, ByVal vDiscAmount As Decimal, ByVal vNettPrice As Decimal, _
    ByVal vSubTotal As Decimal, ByVal vPrimaryDisc1 As Double, ByVal vDiscString1 As String, ByVal vDiscAmount1 As Decimal, _
    ByVal vPrimaryDisc2 As Double, ByVal vDiscString2 As String, ByVal vDiscAmount2 As Decimal, _
    ByVal vPrimaryDisc3 As Double, ByVal vDiscString3 As String, ByVal vDiscAmount3 As Decimal, _
    ByVal vPrimaryLocation As Double, ByVal vBatchNumber As String, ByVal vExpiryDate As Date, _
    ByVal vReason As String, ByVal vPrimaryUnit As Double, ByVal vPrimaryUser As Double, ByVal vPrimaryPOS As Double, _
    ByVal vPrimarySession As Double, ByVal vPrimaryProduct As Double, ByVal i As FrmTrans_POS) As Boolean
        Dim DTAItem As New DseTrans_POS_VoidTableAdapters.tbltrans_session_cabang_itemvoid1TableAdapter
        Dim ObjDiskon As New ClsMast_Diskon
        Dim ObjCD As New ClsCekData
        Dim ObjSession As New ClsTrans_Session
        Try
            DTAItem.Connection.ConnectionString = Me.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If String.IsNullOrEmpty(vPrimaryDisc1) = True Then
                vPrimaryDisc1 = "-1"
            End If
            If String.IsNullOrEmpty(vPrimaryDisc2) = True Then
                vPrimaryDisc2 = "-1"
            End If
            If String.IsNullOrEmpty(vPrimaryDisc3) = True Then
                vPrimaryDisc3 = "-1"
            End If
            If String.IsNullOrEmpty(vDiscAmount1) = True Then
                vDiscAmount1 = "0"
            End If
            If String.IsNullOrEmpty(vDiscAmount2) = True Then
                vDiscAmount2 = "0"
            End If
            If String.IsNullOrEmpty(vDiscAmount3) = True Then
                vDiscAmount3 = "0"
            End If
            If String.IsNullOrEmpty(vDiscString1) = True Then
                vDiscString1 = ""
            End If
            If String.IsNullOrEmpty(vDiscString2) = True Then
                vDiscString2 = ""
            End If
            If String.IsNullOrEmpty(vDiscString3) = True Then
                vDiscString3 = ""
            End If
            If String.IsNullOrEmpty(vExpiryDate) = True Then
                vExpiryDate = "1950/01/01"
            End If
            If String.IsNullOrEmpty(vPrimaryLocation) = True Then
                vPrimaryLocation = "-1"
            End If
            If String.IsNullOrEmpty(vReason) = True Then
                vReason = ""
            End If

            'lakukan perhitung ulang 
            vDiscString = ObjDiskon.Get_Disc_String(ObjSession.GetPrimaryBranch(vPrimarySession), vPrimaryProduct, _
                         i.TxtPrimaryMember.Text, i.dtpTanggalTransaksi.Value, vQty, i.TBLDiskon, 0)
            vDiscAmount = ObjCD.Get_Disc_Amount(vDiscString, _
                         CDec(vQty * vUnitPrice), i.DefPrec, i.RoundingTipe, i.RoundingValue, ObjCD)
            vNettPrice = ObjCD.DigitGrouping(vUnitPrice - vDiscAmount, i.DefPrec)
            vSubTotal = ObjCD.DigitGrouping(vQty * vNettPrice, i.DefPrec)
            vDiscAmount1 = ObjCD.Get_Disc_Amount(vDiscString1, _
                         CDec(vQty * vUnitPrice), i.DefPrec, i.RoundingTipe, i.RoundingValue, ObjCD)
            vDiscAmount2 = ObjCD.Get_Disc_Amount(vDiscString2, _
                         CDec(vQty * vUnitPrice), i.DefPrec, i.RoundingTipe, i.RoundingValue, ObjCD)
            vDiscAmount3 = ObjCD.Get_Disc_Amount(vDiscString3, _
                         CDec(vQty * vUnitPrice), i.DefPrec, i.RoundingTipe, i.RoundingValue, ObjCD)

            '-----
            DTAItem.InsertQuery(vPrimarySession, vPrimaryProduct, vPrimaryPOS, vQty, vUnitPrice, vDiscString, vDiscAmount, vNettPrice, _
            vSubTotal, vPrimaryDisc1, vDiscString1, vDiscAmount1, vPrimaryDisc2, vDiscString2, vDiscAmount2, vPrimaryDisc3, vDiscString3, _
            vDiscAmount3, vPrimaryLocation, vBatchNumber, vExpiryDate, vPrimaryUnit, vPrimaryUser, vReason)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            ObjDiskon = Nothing
            ObjCD = Nothing
            ObjSession = Nothing
        End Try
    End Function

    Public Function UpdateItemViod(ByVal vQty As Decimal, _
    ByVal vUnitPrice As Decimal, ByVal vDiscString As String, ByVal vDiscAmount As Decimal, ByVal vNettPrice As Decimal, _
    ByVal vSubTotal As Decimal, ByVal vPrimaryDisc1 As Double, ByVal vDiscString1 As String, ByVal vDiscAmount1 As Decimal, _
    ByVal vPrimaryDisc2 As Double, ByVal vDiscString2 As String, ByVal vDiscAmount2 As Decimal, _
    ByVal vPrimaryDisc3 As Double, ByVal vDiscString3 As String, ByVal vDiscAmount3 As Decimal, _
    ByVal vPrimaryLocation As Double, ByVal vBatchNumber As String, ByVal vExpiryDate As Date, _
    ByVal vReason As String, ByVal vPrimaryUnit As Double, ByVal vPrimaryUser As Double, ByVal vPrimaryPOS As Double, _
    ByVal vPrimarySession As Double, ByVal vPrimaryProduct As Double, ByVal i As FrmTrans_POS) As Boolean
        Dim DTAItem As New DseTrans_POS_VoidTableAdapters.tbltrans_session_cabang_itemvoid1TableAdapter
        Dim ObjDiskon As New ClsMast_Diskon
        Dim ObjCD As New ClsCekData
        Dim ObjSession As New ClsTrans_Session
        Try
            DTAItem.Connection.ConnectionString = Me.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If String.IsNullOrEmpty(vPrimaryDisc1) = True Then
                vPrimaryDisc1 = "-1"
            End If
            If String.IsNullOrEmpty(vPrimaryDisc2) = True Then
                vPrimaryDisc2 = "-1"
            End If
            If String.IsNullOrEmpty(vPrimaryDisc3) = True Then
                vPrimaryDisc3 = "-1"
            End If
            If String.IsNullOrEmpty(vDiscAmount1) = True Then
                vDiscAmount1 = "0"
            End If
            If String.IsNullOrEmpty(vDiscAmount2) = True Then
                vDiscAmount2 = "0"
            End If
            If String.IsNullOrEmpty(vDiscAmount3) = True Then
                vDiscAmount3 = "0"
            End If
            If String.IsNullOrEmpty(vDiscString1) = True Then
                vDiscString1 = ""
            End If
            If String.IsNullOrEmpty(vDiscString2) = True Then
                vDiscString2 = ""
            End If
            If String.IsNullOrEmpty(vDiscString3) = True Then
                vDiscString3 = ""
            End If
            If String.IsNullOrEmpty(vExpiryDate) = True Then
                vExpiryDate = "1950/01/01"
            End If
            If String.IsNullOrEmpty(vPrimaryLocation) = True Then
                vPrimaryLocation = "-1"
            End If
            If String.IsNullOrEmpty(vReason) = True Then
                vReason = ""
            End If

            vQty += Me.GetQty(vPrimaryPOS, vPrimaryProduct, vPrimarySession, vPrimaryUser)
            'lakukan perhitung ulang 
            vDiscString = ObjDiskon.Get_Disc_String(ObjSession.GetPrimaryBranch(vPrimarySession), vPrimaryProduct, _
                         i.TxtPrimaryMember.Text, i.dtpTanggalTransaksi.Value, vQty, i.TBLDiskon, 0)
            vDiscAmount = ObjCD.Get_Disc_Amount(vDiscString, _
                         CDec(vQty * vUnitPrice), i.DefPrec, i.RoundingTipe, i.RoundingValue, ObjCD)
            vNettPrice = ObjCD.DigitGrouping(vUnitPrice - vDiscAmount, i.DefPrec)
            vSubTotal = ObjCD.DigitGrouping(vQty * vNettPrice, i.DefPrec)
            vDiscAmount1 = ObjCD.Get_Disc_Amount(vDiscString1, _
                         CDec(vQty * vUnitPrice), i.DefPrec, i.RoundingTipe, i.RoundingValue, ObjCD)
            vDiscAmount2 = ObjCD.Get_Disc_Amount(vDiscString2, _
                         CDec(vQty * vUnitPrice), i.DefPrec, i.RoundingTipe, i.RoundingValue, ObjCD)
            vDiscAmount3 = ObjCD.Get_Disc_Amount(vDiscString3, _
                         CDec(vQty * vUnitPrice), i.DefPrec, i.RoundingTipe, i.RoundingValue, ObjCD)

            '-----
            DTAItem.UpdateQuery(vQty, vUnitPrice, vDiscString, vDiscAmount, vNettPrice, vSubTotal, vPrimaryDisc1, vDiscString1, vDiscAmount1, _
            vPrimaryDisc2, vDiscString2, vDiscAmount2, vPrimaryDisc3, vDiscString3, vDiscAmount3, vPrimaryLocation, vBatchNumber, _
            vExpiryDate, vReason, vPrimaryUnit, vPrimarySession, vPrimaryProduct, vPrimaryPOS, vPrimaryUser)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            ObjDiskon = Nothing
            ObjCD = Nothing
            ObjSession = Nothing
        End Try
    End Function

    Public Function HapusItemViod(ByVal vPrimaryPOS As Double, _
    ByVal vPrimarySession As Double, ByVal vPrimaryProduct As Double) As Boolean
        Dim DTAItem As New DseTrans_POS_VoidTableAdapters.tbltrans_session_cabang_itemvoid1TableAdapter
        Try
            DTAItem.Connection.ConnectionString = Me.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DTAItem.DeleteQuery(vPrimarySession, vPrimaryProduct, vPrimaryPOS)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function
End Class
