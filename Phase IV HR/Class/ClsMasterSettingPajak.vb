Option Explicit On
Option Strict Off

Imports System.Windows.Forms

Public Class ClsMasterSettingPajak
    Public Enum EnumCriteria As Integer
        TK0
        TK1
        TK2
        TK3
        K0
        K1
        K2
        K3
    End Enum

    Public Function GetDesc(ByVal vEnumCriteria As EnumCriteria) As String
        Try
            If vEnumCriteria = EnumCriteria.K0 Then
                Return "K0"
            ElseIf vEnumCriteria = EnumCriteria.K1 Then
                Return "K1"
            ElseIf vEnumCriteria = EnumCriteria.K2 Then
                Return "K2"
            ElseIf vEnumCriteria = EnumCriteria.K3 Then
                Return "K3"
            ElseIf vEnumCriteria = EnumCriteria.TK0 Then
                Return "TK0"
            ElseIf vEnumCriteria = EnumCriteria.TK1 Then
                Return "TK1"
            ElseIf vEnumCriteria = EnumCriteria.TK2 Then
                Return "TK2"
            ElseIf vEnumCriteria = EnumCriteria.TK3 Then
                Return "TK3"
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetPTKPPerTahun(ByVal vEnumCriteria As EnumCriteria) As Decimal
        Dim DTA As New DseMasterSettingPajakTableAdapters.tblmaster_ptkpTableAdapter
        Dim Dummy As System.Nullable(Of Decimal)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPTKPTahunan(vEnumCriteria)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlPTKP() As Double
        Dim DTA As New DseMasterSettingPajakTableAdapters.tblmaster_ptkpTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlPTKP
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda_PTKP() As Boolean
        Try
            If Me.GetJmlPTKP() > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsiDGVItem(ByVal vDGV_PTKP As DataGridView, ByVal vDGV_PPh As DataGridView, _
    ByVal Frm As FrmMasterSettingPajak, ByVal vDGV_PPh_Final As DataGridView) As Boolean
        Dim DTA_PTKP As New DseMasterSettingPajakTableAdapters.tblmaster_ptkpTableAdapter
        Dim TBL_PTKP As New DseMasterSettingPajak.tblmaster_ptkpDataTable
        Dim Row_PTKP As DseMasterSettingPajak.tblmaster_ptkpRow
        Dim DTA_PPh As New DseMasterSettingPajakTableAdapters.tblmaster_tarif_pphTableAdapter
        Dim TBL_PPh As New DseMasterSettingPajak.tblmaster_tarif_pphDataTable
        Dim Row_PPh As DseMasterSettingPajak.tblmaster_tarif_pphRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCurr As New Core_Module.clsMasterCurrency
        Dim DefPrec As Double = ObjCurr.GetDecimal_DefaultCurrency
        Dim DefPrec_Persen As Double = 2
        Dim rowDGV As DataGridViewRow
        Dim DTA_PPh_Final As New DseMasterSettingPajakTableAdapters.tblmaster_tarif_pph_final_pesangonTableAdapter
        Dim TBL_PPh_Final As New DseMasterSettingPajak.tblmaster_tarif_pph_final_pesangonDataTable
        Dim Row_PPh_Final As DseMasterSettingPajak.tblmaster_tarif_pph_final_pesangonRow

        Try
            DTA_PPh_Final.Connection.ConnectionString = Core_Module.GetConstring
            DTA_PTKP.Connection.ConnectionString = DTA_PPh_Final.Connection.ConnectionString
            DTA_PPh.Connection.ConnectionString = DTA_PPh_Final.Connection.ConnectionString

            'kosongkan dgv
            vDGV_PPh.AllowUserToAddRows = False
            vDGV_PPh_Final.AllowUserToAddRows = False
            vDGV_PTKP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV_PPh.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV_PPh_Final.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV_PTKP.Rows.Clear()
            vDGV_PPh.Rows.Clear()
            vDGV_PPh_Final.Rows.Clear()

            'isikan(tabel)
            If DTA_PTKP.Fill(TBL_PTKP) > 0 Then
                For Each Row_PTKP In TBL_PTKP.Rows
                    With Row_PTKP
                        vDGV_PTKP.Rows.Add("")
                        rowDGV = vDGV_PTKP.Rows(vDGV_PTKP.Rows.Count - 1)
                        If .enum_golongan_pajak = ClsMasterSettingPajak.EnumCriteria.TK0 Then
                            rowDGV.Cells(Frm.ColPTKP_Golongan.Index).Value = "TK-0"
                        End If
                        If .enum_golongan_pajak = ClsMasterSettingPajak.EnumCriteria.TK1 Then
                            rowDGV.Cells(Frm.ColPTKP_Golongan.Index).Value = "TK-1"
                        End If
                        If .enum_golongan_pajak = ClsMasterSettingPajak.EnumCriteria.TK2 Then
                            rowDGV.Cells(Frm.ColPTKP_Golongan.Index).Value = "TK-2"
                        End If
                        If .enum_golongan_pajak = ClsMasterSettingPajak.EnumCriteria.TK3 Then
                            rowDGV.Cells(Frm.ColPTKP_Golongan.Index).Value = "TK-3"
                        End If
                        If .enum_golongan_pajak = ClsMasterSettingPajak.EnumCriteria.K0 Then
                            rowDGV.Cells(Frm.ColPTKP_Golongan.Index).Value = "K-0"
                        End If
                        If .enum_golongan_pajak = ClsMasterSettingPajak.EnumCriteria.K1 Then
                            rowDGV.Cells(Frm.ColPTKP_Golongan.Index).Value = "K-1"
                        End If
                        If .enum_golongan_pajak = ClsMasterSettingPajak.EnumCriteria.K2 Then
                            rowDGV.Cells(Frm.ColPTKP_Golongan.Index).Value = "K-2"
                        End If
                        If .enum_golongan_pajak = ClsMasterSettingPajak.EnumCriteria.K3 Then
                            rowDGV.Cells(Frm.ColPTKP_Golongan.Index).Value = "K-3"
                        End If
                        rowDGV.Cells(Frm.ColPTKP_Enum.Index).Value = .enum_golongan_pajak
                        rowDGV.Cells(Frm.ColPTKP_PTKP_Tahunan.Index).Value = ObjCD.DigitGrouping(.PTKP_Tahunan, DefPrec)
                        rowDGV.Cells(Frm.ColPTKP_Remarks.Index).Value = .remarks
                        rowDGV.Cells(Frm.ColPTKP_Primary.Index).Value = .primary
                    End With
                Next
            End If

            If DTA_PPh.Fill(TBL_PPh) > 0 Then
                For Each Row_PPh In TBL_PPh.Rows
                    With Row_PPh
                        vDGV_PPh.Rows.Add("")
                        rowDGV = vDGV_PPh.Rows(vDGV_PPh.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColRate_Maximum.Index).Value = ObjCD.DigitGrouping(.maximum_income, DefPrec)
                        rowDGV.Cells(Frm.ColRate_Tarif.Index).Value = ObjCD.DigitGrouping(.tarif, DefPrec_Persen)
                        rowDGV.Cells(Frm.ColRate_Remarks.Index).Value = .remarks
                        rowDGV.Cells(Frm.ColRate_Primary.Index).Value = .primary
                    End With
                Next
            End If

            If DTA_PPh_Final.Fill(TBL_PPh_Final) > 0 Then
                For Each Row_PPh_Final In TBL_PPh_Final.Rows
                    With Row_PPh_Final
                        vDGV_PPh_Final.Rows.Add("")
                        rowDGV = vDGV_PPh_Final.Rows(vDGV_PPh_Final.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColRateFinal_Maksimum.Index).Value = ObjCD.DigitGrouping(.maximum_income, DefPrec)
                        rowDGV.Cells(Frm.ColRateFinal_Tarif.Index).Value = ObjCD.DigitGrouping(.tarif, DefPrec_Persen)
                        rowDGV.Cells(Frm.ColRateFinal_Remarks.Index).Value = .remarks
                        rowDGV.Cells(Frm.ColRateFinal_Primary.Index).Value = .primary
                    End With
                Next
            End If

            'set nomer
            ''ObjCD.SetNumber(vDGVItem)
            vDGV_PTKP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV_PPh.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV_PPh.AllowUserToAddRows = True
            vDGV_PPh_Final.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV_PPh_Final.AllowUserToAddRows = True

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_PTKP.Dispose()
            TBL_PTKP.Dispose()
            DTA_PPh.Dispose()
            TBL_PPh.Dispose()
            ObjCD = Nothing
            ObjCurr = Nothing
            DTA_PPh_Final.Dispose()
            TBL_PPh_Final.Dispose()
        End Try
    End Function

    Public Function IsiDGVItem_Default(ByVal vDGV_PTKP As DataGridView, ByVal vDGV_PPh As DataGridView, _
    ByVal Frm As FrmMasterSettingPajak, ByVal vDGV_PPh_Final As DataGridView) As Boolean
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCurr As New Core_Module.clsMasterCurrency
        Dim DefPrec As Double = ObjCurr.GetDecimal_DefaultCurrency
        Dim DefPrec_Persen As Double = 2
        Dim rowDGV As DataGridViewRow
        Dim i As Integer = 0

        Try
            'kosongkan dgv
            vDGV_PPh.AllowUserToAddRows = False
            vDGV_PPh_Final.AllowUserToAddRows = False
            vDGV_PTKP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV_PPh.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV_PPh_Final.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV_PTKP.Rows.Clear()
            vDGV_PPh.Rows.Clear()
            vDGV_PPh_Final.Rows.Clear()

            i = 0
            While i <= 7
                vDGV_PTKP.Rows.Add("")
                rowDGV = vDGV_PTKP.Rows(vDGV_PTKP.Rows.Count - 1)
                If i = 0 Then
                    rowDGV.Cells(Frm.ColPTKP_Golongan.Index).Value = "TK-0"
                End If
                If i = 1 Then
                    rowDGV.Cells(Frm.ColPTKP_Golongan.Index).Value = "TK-1"
                End If
                If i = 2 Then
                    rowDGV.Cells(Frm.ColPTKP_Golongan.Index).Value = "TK-2"
                End If
                If i = 3 Then
                    rowDGV.Cells(Frm.ColPTKP_Golongan.Index).Value = "TK-3"
                End If
                If i = 4 Then
                    rowDGV.Cells(Frm.ColPTKP_Golongan.Index).Value = "K-0"
                End If
                If i = 5 Then
                    rowDGV.Cells(Frm.ColPTKP_Golongan.Index).Value = "K-1"
                End If
                If i = 6 Then
                    rowDGV.Cells(Frm.ColPTKP_Golongan.Index).Value = "K-2"
                End If
                If i = 7 Then
                    rowDGV.Cells(Frm.ColPTKP_Golongan.Index).Value = "K-3"
                End If
                rowDGV.Cells(Frm.ColPTKP_Enum.Index).Value = i
                rowDGV.Cells(Frm.ColPTKP_PTKP_Tahunan.Index).Value = ObjCD.DigitGrouping(0, DefPrec)
                rowDGV.Cells(Frm.ColPTKP_Remarks.Index).Value = ""
                rowDGV.Cells(Frm.ColPTKP_Primary.Index).Value = -1
                i += 1
            End While

            i = 0
            While i <= 3
                vDGV_PPh.Rows.Add("")
                rowDGV = vDGV_PPh.Rows(vDGV_PPh.Rows.Count - 1)
                If i = 0 Then
                    rowDGV.Cells(Frm.ColRate_Maximum.Index).Value = ObjCD.DigitGrouping(50000000, DefPrec)
                    rowDGV.Cells(Frm.ColRate_Tarif.Index).Value = ObjCD.DigitGrouping(5, DefPrec_Persen)
                End If
                If i = 1 Then
                    rowDGV.Cells(Frm.ColRate_Maximum.Index).Value = ObjCD.DigitGrouping(250000000, DefPrec)
                    rowDGV.Cells(Frm.ColRate_Tarif.Index).Value = ObjCD.DigitGrouping(15, DefPrec_Persen)
                End If
                If i = 2 Then
                    rowDGV.Cells(Frm.ColRate_Maximum.Index).Value = ObjCD.DigitGrouping(500000000, DefPrec)
                    rowDGV.Cells(Frm.ColRate_Tarif.Index).Value = ObjCD.DigitGrouping(25, DefPrec_Persen)
                End If
                If i = 3 Then
                    rowDGV.Cells(Frm.ColRate_Maximum.Index).Value = ObjCD.DigitGrouping(999999999999, DefPrec)
                    rowDGV.Cells(Frm.ColRate_Tarif.Index).Value = ObjCD.DigitGrouping(30, DefPrec_Persen)
                End If
                rowDGV.Cells(Frm.ColRate_Remarks.Index).Value = ""
                rowDGV.Cells(Frm.ColRate_Primary.Index).Value = ""
                i += 1
            End While

            i = 0
            While i <= 3
                vDGV_PPh_Final.Rows.Add("")
                rowDGV = vDGV_PPh_Final.Rows(vDGV_PPh_Final.Rows.Count - 1)
                If i = 0 Then
                    rowDGV.Cells(Frm.ColRateFinal_Maksimum.Index).Value = ObjCD.DigitGrouping(50000000, DefPrec)
                    rowDGV.Cells(Frm.ColRateFinal_Tarif.Index).Value = ObjCD.DigitGrouping(0, DefPrec_Persen)
                End If
                If i = 1 Then
                    rowDGV.Cells(Frm.ColRateFinal_Maksimum.Index).Value = ObjCD.DigitGrouping(100000000, DefPrec)
                    rowDGV.Cells(Frm.ColRateFinal_Tarif.Index).Value = ObjCD.DigitGrouping(5, DefPrec_Persen)
                End If
                If i = 2 Then
                    rowDGV.Cells(Frm.ColRateFinal_Maksimum.Index).Value = ObjCD.DigitGrouping(500000000, DefPrec)
                    rowDGV.Cells(Frm.ColRateFinal_Tarif.Index).Value = ObjCD.DigitGrouping(15, DefPrec_Persen)
                End If
                If i = 3 Then
                    rowDGV.Cells(Frm.ColRateFinal_Maksimum.Index).Value = ObjCD.DigitGrouping(999999999999, DefPrec)
                    rowDGV.Cells(Frm.ColRateFinal_Tarif.Index).Value = ObjCD.DigitGrouping(25, DefPrec_Persen)
                End If
                rowDGV.Cells(Frm.ColRateFinal_Remarks.Index).Value = ""
                rowDGV.Cells(Frm.ColRateFinal_Primary.Index).Value = ""
                i += 1
            End While

            'set nomer
            ''ObjCD.SetNumber(vDGVItem)
            vDGV_PTKP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV_PPh.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV_PPh.AllowUserToAddRows = True
            vDGV_PPh_Final.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGV_PPh_Final.AllowUserToAddRows = True

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
            ObjCurr = Nothing
        End Try
    End Function

    Public Function Simpan(ByVal vDGVItem_PTKP As DataGridView, ByVal vDGVItem_PPh As DataGridView, ByVal i As FrmMasterSettingPajak, _
    ByVal vDGVItem_PPh_Final As DataGridView) As Boolean
        Dim DTA_PTKP As New DseMasterSettingPajakTableAdapters.tblmaster_ptkpTableAdapter
        Dim DTA_PPh As New DseMasterSettingPajakTableAdapters.tblmaster_tarif_pphTableAdapter
        Dim DTA_PPh_Final As New DseMasterSettingPajakTableAdapters.tblmaster_tarif_pph_final_pesangonTableAdapter
        Dim RowDGV_PTKP As DataGridViewRow
        Dim RowDGV_PPh As DataGridViewRow
        Dim RowDGV_PPh_Final As DataGridViewRow

        Try
            vDGVItem_PPh.AllowUserToAddRows = False
            vDGVItem_PPh_Final.AllowUserToAddRows = False

            'Setting Connection
            DTA_PTKP.Connection.ConnectionString = Core_Module.GetConstring
            DTA_PPh.Connection.ConnectionString = Core_Module.GetConstring
            DTA_PPh_Final.Connection.ConnectionString = Core_Module.GetConstring

            'PTKP
            If Me.ApakahSudahAda_PTKP() = True Then
                'Update data
                If vDGVItem_PTKP.Rows.Count > 0 Then
                    For Each RowDGV_PTKP In vDGVItem_PTKP.Rows
                        DTA_PTKP.Update(RowDGV_PTKP.Cells(i.ColPTKP_Enum.Index).Value, RowDGV_PTKP.Cells(i.ColPTKP_PTKP_Tahunan.Index).Value, _
                         RowDGV_PTKP.Cells(i.ColPTKP_Remarks.Index).Value, RowDGV_PTKP.Cells(i.ColPTKP_Primary.Index).Value)
                    Next
                End If
            Else
                'Insert data
                If vDGVItem_PTKP.Rows.Count > 0 Then
                    For Each RowDGV_PTKP In vDGVItem_PTKP.Rows
                        DTA_PTKP.Insert(RowDGV_PTKP.Cells(i.ColPTKP_Enum.Index).Value, RowDGV_PTKP.Cells(i.ColPTKP_PTKP_Tahunan.Index).Value, _
                        RowDGV_PTKP.Cells(i.ColPTKP_Remarks.Index).Value)
                    Next
                End If
            End If
            'PPh
            If vDGVItem_PPh.Rows.Count > 0 Then
                DTA_PPh.DeleteQuery_PPhAllData()
                For Each RowDGV_PPh In vDGVItem_PPh.Rows
                    If String.IsNullOrEmpty(RowDGV_PPh.Cells(i.ColRate_Remarks.Index).Value) = True Then
                        RowDGV_PPh.Cells(i.ColRate_Remarks.Index).Value = ""
                    End If
                    DTA_PPh.Insert(RowDGV_PPh.Cells(i.ColRate_Maximum.Index).Value, RowDGV_PPh.Cells(i.ColRate_Tarif.Index).Value, _
                    RowDGV_PPh.Cells(i.ColRate_Remarks.Index).Value)
                Next
            End If

            'PPh Final
            If vDGVItem_PPh_Final.Rows.Count > 0 Then
                DTA_PPh_Final.DeleteQuery()
                For Each RowDGV_PPh_Final In vDGVItem_PPh_Final.Rows
                    If String.IsNullOrEmpty(RowDGV_PPh_Final.Cells(i.ColRateFinal_Remarks.Index).Value) = True Then
                        RowDGV_PPh_Final.Cells(i.ColRateFinal_Remarks.Index).Value = ""
                    End If
                    DTA_PPh_Final.InsertQuery(RowDGV_PPh_Final.Cells(i.ColRateFinal_Maksimum.Index).Value, RowDGV_PPh_Final.Cells(i.ColRateFinal_Tarif.Index).Value, _
                    RowDGV_PPh_Final.Cells(i.ColRateFinal_Remarks.Index).Value)
                Next
            End If

            vDGVItem_PPh.AllowUserToAddRows = True
            vDGVItem_PPh_Final.AllowUserToAddRows = True

            'Load data
            Me.IsiDGVItem(vDGVItem_PTKP, vDGVItem_PPh, i, vDGVItem_PPh_Final)

            'jika sudah sampai sini
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_PTKP.Dispose()
            DTA_PPh.Dispose()
            DTA_PPh_Final.Dispose()
        End Try
    End Function

    Public Function ApakahDataBenar(ByVal vDGVItem_PTKP As DataGridView, ByVal vDGVItem_PPh As DataGridView, _
    ByVal vShowMsg As Boolean, ByVal i As FrmMasterSettingPajak, ByVal vDGVItem_PPh_Final As DataGridView) As Boolean
        Dim ObjCD As New Core_Module.ClsCekData
        Dim RowDGV As DataGridViewRow
        Dim IndexDGV As Integer = -1
        Try
            If vDGVItem_PTKP.Rows.Count > 0 Then
                For Each RowDGV In vDGVItem_PTKP.Rows
                    If RowDGV.Cells(i.ColPTKP_PTKP_Tahunan.Index).Value = False Then
                        IndexDGV = RowDGV.Index
                        If vShowMsg = True Then
                            MsgBox("Please enter positive number in PTKP Tahunan at line " + CStr(IndexDGV + 1) + ".", MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                Next
            Else
                If vShowMsg = True Then
                    MsgBox("You cannot input PTKP less than 8 data.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            vDGVItem_PPh.AllowUserToAddRows = False
            If vDGVItem_PPh.Rows.Count <= 0 Then
                If vShowMsg = True Then
                    MsgBox("Please input rate data.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            vDGVItem_PPh.AllowUserToAddRows = True

            If Me.ApakahDataBenar_RateFinal(vDGVItem_PPh_Final, vShowMsg, i) = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Public Function ApakahDataBenar_RateFinal(ByVal vDGVItem_PPh_Final As DataGridView, _
    ByVal vShowMsg As Boolean, ByVal i As FrmMasterSettingPajak) As Boolean
        Dim ObjCD As New Core_Module.ClsCekData
        Dim IndexDGV As Integer = -1

        Try
            vDGVItem_PPh_Final.AllowUserToAddRows = False
            If vDGVItem_PPh_Final.Rows.Count <= 0 Then
                If vShowMsg = True Then
                    MsgBox("Please input PPH Final income tax data rate data.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            vDGVItem_PPh_Final.AllowUserToAddRows = True
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Function
End Class
