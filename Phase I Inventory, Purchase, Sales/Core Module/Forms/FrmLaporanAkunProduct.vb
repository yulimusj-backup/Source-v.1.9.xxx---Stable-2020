Option Explicit On
Option Strict Off

Public Class FrmLaporanAkunProduct
    Dim DS As New DseLaporanAkunProduct
    Dim report As New CrpLaporanAkunProduk

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmLaporanAkunProduct_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.CboProduct_CodeDesc.SelectedIndex = 0
            LoadLaporan()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkGoods_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkGoods.CheckedChanged
        Try
            LoadLaporan()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkService_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkService.CheckedChanged
        Try
            LoadLaporan()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboProduct_CodeDesc_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboProduct_CodeDesc.SelectedIndexChanged
        Try
            LoadLaporan()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadLaporan()
        Dim DS As New DseLaporanAkunProduct
        Dim DTA As New DseLaporanAkunProductTableAdapters.DataTable1TableAdapter
        Dim TBL As DseLaporanAkunProduct.DataTable1DataTable = DS.DataTable1
        Dim TBL1 As New DseLaporanAkunProduct.DataTable1DataTable
        Dim Row, Row1 As DseLaporanAkunProduct.DataTable1Row

        Dim DV As DataRowView
        Dim paramCbo As Integer
        Dim Filter1 As String = ""
        Dim Filter2 As String = ""
        Dim Goods1 As String = ""
        Dim Goods2 As String = ""
        Dim Service1 As String = ""
        Dim Service2 As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            paramCbo = Me.CboProduct_CodeDesc.SelectedIndex
            If Me.ChkGoods.Checked = True And Me.ChkService.Checked = True Then
                DTA.Fill(TBL)
            ElseIf Me.ChkGoods.Checked = True Then
                DTA.FillBy_GoodsType(TBL)
            ElseIf Me.ChkService.Checked = True Then
                DTA.FillBy_ServiceType(TBL)
            End If

            For Each Row In TBL.Rows
                With Row
                    Filter1 = "[Primary] = " & .Primary & " AND [Type] = " & 0
                    Filter2 = "[Primary] = " & .Primary & " AND [Type] = " & 1
                    Goods1 = "Inventory Account" & vbNewLine & "Sales Account" & vbNewLine & _
                            "Sales Return Account" & vbNewLine & "Cost Of Sales Account" & vbNewLine & "Adjustment Account" _
                            & vbNewLine & "Stock Taking Account"
                    Goods2 = ": " & .Inventory_Account_Code & " - " & .Inventory_Account_Desc & vbNewLine & ": " _
                            & .Sales_Account_Code & " - " & .Sales_Account_Desc & vbNewLine & ": " _
                            & .Sales_Return_Code & " - " & .Sales_Return_Desc & vbNewLine & ": " _
                            & .COGS_Account_Code & " - " & .COGS_Account_Desc & vbNewLine & ": " _
                            & .Adjustment_Account_Code & " - " & .Adjustment_Account_Desc & vbNewLine & ": " _
                            & .Stock_Taking_Account_Code & " - " & .Stock_Taking_Account_Desc
                    Service1 = "Expense Account" & vbNewLine & "Sales Account" & vbNewLine & "Sales Return Account"
                    Service2 = ": " & .Inventory_Account_Code & " - " & .Inventory_Account_Desc & vbNewLine & _
                            ": " & .Sales_Account_Code & " - " & .Sales_Account_Desc & vbNewLine & ": " _
                            & .Sales_Return_Code & " - " & .Sales_Return_Desc

                    If Me.ChkGoods.Checked = True And Me.ChkService.Checked = True Then
                        TBL.DefaultView.RowFilter = Nothing
                        TBL.DefaultView.RowFilter = Filter1
                        For Each DV In TBL.DefaultView
                            .Primary = .Primary
                            .Account_1 = Goods1
                            .Account_2 = Goods2
                        Next
                        TBL.DefaultView.RowFilter = Nothing
                        TBL.DefaultView.RowFilter = Filter2
                        For Each DV In TBL.DefaultView
                            .Primary = .Primary
                            .Account_1 = Service1
                            .Account_2 = Service2
                        Next
                    ElseIf Me.ChkGoods.Checked = True Then
                        .Primary = .Primary
                        .Account_1 = Goods1
                        .Account_2 = Goods2
                    ElseIf Me.ChkService.Checked = True Then
                        .Primary = .Primary
                        .Account_1 = Service1
                        .Account_2 = Service2
                    End If

                    If paramCbo = 0 Then
                        .Product_Code = .Product_Code
                    ElseIf paramCbo = 1 Then
                        .Product_Code = .Description
                    ElseIf paramCbo = 2 Then
                        .Product_Code = .Product_Code & vbNewLine & .Description
                    End If
                End With
            Next

            TBL1.Rows.Clear()
            For Each Row In TBL.Rows
                Row.Product_Code_1 = ""
                Row.Product_Code_2 = ""

                TBL1.DefaultView.RowFilter = "[Account 1] = '" + Row.Account_1 + "' and [Account 2] = '" + Row.Account_2 + "'"
                If TBL1.DefaultView.Count = 0 Then
                    'tambah baris baru
                    Row1 = TBL1.NewDataTable1Row
                    For ij As Integer = 0 To TBL1.Columns.Count - 1
                        Row1.Item(ij) = Row.Item(ij)
                    Next
                    Row1.Primary = TBL1.Rows.Count + 1
                    TBL1.Rows.Add(Row1)
                Else
                    TBL1.DefaultView.Sort = "Primary"
                    Row1 = TBL1.DefaultView(TBL1.DefaultView.Count - 1).Row
                    If Row1.Product_Code_1 <> "" And Row1.Product_Code_2 <> "" Then
                        'tambah baris baru
                        Row1 = TBL1.NewDataTable1Row
                        For ij As Integer = 0 To TBL1.Columns.Count - 1
                            Row1.Item(ij) = Row.Item(ij)
                        Next
                        Row1.Primary = TBL1.Rows.Count + 1
                        TBL1.Rows.Add(Row1)
                    Else
                        'product code 1 & 2
                        If Row1.Product_Code_1 = "" Then
                            Row1.Product_Code_1 = Row.Product_Code
                        Else
                            Row1.Product_Code_2 = Row.Product_Code
                        End If
                    End If
                End If
            Next

            TBL.Rows.Clear()
            For Each Row1 In TBL1.Rows
                Row = TBL.NewDataTable1Row
                For ij As Integer = 0 To TBL.Columns.Count - 1
                    Row.Item(ij) = Row1.Item(ij)
                Next
                TBL.Rows.Add(Row)
            Next

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_Cbo.ParameterFieldName, paramCbo)
            crv1.ReportSource = report
            crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub
End Class