''@rosadinnn
''@20150912 1041

Option Explicit On
Option Strict Off

Public Class FrmMasterProductListReport

    Public UserCode As String = ""
    Public BackColorAll As Color = Color.AliceBlue

    Dim ObjCd As New Core_Module.ClsCekData
    Dim RowDgv As DataGridViewRow
    Dim ListBrand As New List(Of Double)
    Dim ListSubBrand As New List(Of Double)
    Dim ListCategory As New List(Of Double)
    Dim ListProduct As New List(Of Double)
    Dim DgvCurrent As DataGridView
    Dim IdxCurrent As Integer
    Dim Buffer As String = ""

    Enum EnumOption As Integer
        PG
        B
        SB
        C
    End Enum

    Private Sub FrmProductListReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.DgvProductGroup.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.DgvBrand.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.DgvSubBrand.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.DgvCategory.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.DgvProduct.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            FrmMasterProductListReport.CheckForIllegalCrossThreadCalls = False
            LoadAllProductGroup()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub UnShow(ByVal vPrimary As Double, ByVal vUnShowFrom As EnumOption)
        Try
            Me.ListBrand.Clear()
            Me.ListSubBrand.Clear()
            Me.ListCategory.Clear()
            Me.ListProduct.Clear()

            If vUnShowFrom = EnumOption.PG Then
                Me.GetListBrand(vPrimary)
                Me.GetListSubBrand(Me.ListBrand)
                Me.GetListCategory(Me.ListSubBrand)
                Me.GetListProduct(Me.ListCategory)
            ElseIf vUnShowFrom = EnumOption.B Then
                Me.ListBrand.Add(vPrimary)
                Me.GetListSubBrand(Me.ListBrand)
                Me.GetListCategory(Me.ListSubBrand)
                Me.GetListProduct(Me.ListCategory)
            ElseIf vUnShowFrom = EnumOption.SB Then
                Me.ListSubBrand.Add(vPrimary)
                Me.GetListCategory(Me.ListSubBrand)
                Me.GetListProduct(Me.ListCategory)
            ElseIf vUnShowFrom = EnumOption.C Then
                Me.ListCategory.Add(vPrimary)
                Me.GetListProduct(Me.ListCategory)
            End If

            ''--remove brand
            If vUnShowFrom = EnumOption.PG Then
                Me.DgvBrand.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvBrand.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                For i As Integer = Me.DgvBrand.Rows.Count - 1 To 0 Step -1
                    If Me.ListBrand.Contains(Me.DgvBrand.Item(Me.ColPrimary_B.Index, i).Value) = True Then
                        Me.DgvBrand.Rows.RemoveAt(i)
                    End If
                Next
                Me.DgvBrand.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.DgvBrand.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                Me.ObjCd.SetNumber(Me.DgvBrand)
            End If

            ''--remove sub brand
            If vUnShowFrom = EnumOption.PG Or vUnShowFrom = EnumOption.B Then
                Me.DgvSubBrand.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvSubBrand.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                For i As Integer = Me.DgvSubBrand.Rows.Count - 1 To 0 Step -1
                    If Me.ListSubBrand.Contains(Me.DgvSubBrand.Item(Me.ColPrimary_SB.Index, i).Value) = True Then
                        Me.DgvSubBrand.Rows.RemoveAt(i)
                    End If
                Next
                Me.DgvSubBrand.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.DgvSubBrand.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                Me.ObjCd.SetNumber(Me.DgvSubBrand)
            End If

            ''--remove category
            If vUnShowFrom = EnumOption.PG Or vUnShowFrom = EnumOption.B Or vUnShowFrom = EnumOption.SB Then
                Me.DgvCategory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvCategory.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                For i As Integer = Me.DgvCategory.Rows.Count - 1 To 0 Step -1
                    If Me.ListCategory.Contains(Me.DgvCategory.Item(Me.ColPrimary_C.Index, i).Value) = True Then
                        Me.DgvCategory.Rows.RemoveAt(i)
                    End If
                Next
                Me.DgvCategory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.DgvCategory.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                Me.ObjCd.SetNumber(Me.DgvCategory)
            End If

            ''--remove product
            If vUnShowFrom = EnumOption.PG Or vUnShowFrom = EnumOption.B Or vUnShowFrom = EnumOption.SB Or vUnShowFrom = EnumOption.C Then
                Me.DgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvProduct.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                For i As Integer = Me.DgvProduct.Rows.Count - 1 To 0 Step -1
                    If Me.ListProduct.Contains(Me.DgvProduct.Item(Me.ColPrimary_P.Index, i).Value) = True Then
                        Me.DgvProduct.Rows.RemoveAt(i)
                    End If
                Next
                Me.DgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.DgvProduct.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                Me.ObjCd.SetNumber(Me.DgvProduct)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub Show_(ByVal vPrimary As Double, ByVal vShowFrom As EnumOption, ByVal vClear As Boolean)
        Try
            Me.ListBrand.Clear()
            Me.ListSubBrand.Clear()
            Me.ListCategory.Clear()
            Me.ListProduct.Clear()

            If vShowFrom = EnumOption.PG Then
                Me.GetListBrand(vPrimary)
                'Me.GetListSubBrand(Me.ListBrand)
                'Me.GetListCategory(Me.ListSubBrand)
                'Me.GetListProduct(Me.ListCategory)
            ElseIf vShowFrom = EnumOption.B Then
                Me.ListBrand.Add(vPrimary)
                Me.GetListSubBrand(Me.ListBrand)
                'Me.GetListCategory(Me.ListSubBrand)
                'Me.GetListProduct(Me.ListCategory)
            ElseIf vShowFrom = EnumOption.SB Then
                Me.ListSubBrand.Add(vPrimary)
                Me.GetListCategory(Me.ListSubBrand)
                'Me.GetListProduct(Me.ListCategory)
            ElseIf vShowFrom = EnumOption.C Then
                Me.ListCategory.Add(vPrimary)
                Me.GetListProduct(Me.ListCategory)
            End If

            ''--add brand
            If vShowFrom = EnumOption.PG Then
                Me.LoadAllBrand(vPrimary, vClear)
            End If

            ''--add sub brand
            If vShowFrom = EnumOption.B Then
                For Each d As Double In Me.ListBrand
                    Me.LoadAllSubBrand(d, vClear)
                Next
            End If

            ''--add category
            If vShowFrom = EnumOption.SB Then
                For Each d As Double In Me.ListSubBrand
                    Me.LoadAllCategory(d, vClear)
                Next
            End If

            ''--add product
            If vShowFrom = EnumOption.C Then
                For Each d As Double In Me.ListCategory
                    Me.LoadAllProduct(d, vClear)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

#Region "load all data"
    Private Sub LoadAllProductGroup()
        Dim DtaPG As New DseProductListReportTableAdapters.tblmaster_product_groupTableAdapter
        Dim TblPG As New DseProductListReport.tblmaster_product_groupDataTable
        Dim RowPG As DseProductListReport.tblmaster_product_groupRow
        Try
            DtaPG.Connection.ConnectionString = Core_Module.GetConstring
            If DtaPG.Fill(TblPG) > 0 Then
                Me.DgvProductGroup.Rows.Clear()
                Me.DgvProductGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvProductGroup.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                For Each RowPG In TblPG
                    Me.DgvProductGroup.Rows.Add()
                    RowDgv = Me.DgvProductGroup.Rows(Me.DgvProductGroup.Rows.Count - 1)
                    RowDgv.Cells(Me.ColNo_PG.Index).Value = ""
                    RowDgv.Cells(Me.ColCode_PG.Index).Value = RowPG.Group_Code
                    RowDgv.Cells(Me.ColDesc_PG.Index).Value = RowPG.Description
                    RowDgv.Cells(Me.ColPrint_PG.Index).Value = Boolean.FalseString
                    RowDgv.Cells(Me.ColPrimary_PG.Index).Value = RowPG.Primary
                Next
                Me.DgvProductGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvProductGroup.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.ObjCd.SetNumber(Me.DgvProductGroup)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub LoadAllBrand(ByVal vPrimary As Double, ByVal vClear As Boolean)
        Dim DtaB As New DseProductListReportTableAdapters.tblmaster_brandTableAdapter
        Dim TblB As New DseProductListReport.tblmaster_brandDataTable
        Dim RowB As DseProductListReport.tblmaster_brandRow
        Dim ListPrimary As New List(Of Double)
        Try
            DtaB.Connection.ConnectionString = Core_Module.GetConstring
            If vPrimary = -1 Then
                DtaB.Fill(TblB)
            Else
                DtaB.FillBy_PrimaryPG(TblB, vPrimary)
            End If
            If TblB.Count > 0 Then
                If vClear = True Then
                    Me.DgvBrand.Rows.Clear()
                Else
                    For Each Me.RowDgv In Me.DgvBrand.Rows
                        ListPrimary.Add(Me.RowDgv.Cells(Me.ColPrimary_B.Index).Value)
                    Next
                End If
                For Each RowB In TblB
                    If ListPrimary.Contains(RowB.Primary) = False Then
                        Me.DgvBrand.Rows.Add()
                        RowDgv = Me.DgvBrand.Rows(Me.DgvBrand.Rows.Count - 1)
                        RowDgv.Cells(Me.ColNo_B.Index).Value = ""
                        RowDgv.Cells(Me.ColCode_B.Index).Value = RowB.Group_Code
                        RowDgv.Cells(Me.ColDesc_B.Index).Value = RowB.Description
                        RowDgv.Cells(Me.ColPrint_B.Index).Value = Boolean.FalseString
                        RowDgv.Cells(Me.ColPrimary_B.Index).Value = RowB.Primary
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub LoadAllSubBrand(ByVal vPrimary As Double, ByVal vClear As Boolean)
        Dim DtaSB As New DseProductListReportTableAdapters.tblmaster_subbrandTableAdapter
        Dim TblSB As New DseProductListReport.tblmaster_subbrandDataTable
        Dim RowSB As DseProductListReport.tblmaster_subbrandRow
        Dim ListPrimary As New List(Of Double)
        Try
            DtaSB.Connection.ConnectionString = Core_Module.GetConstring
            If vPrimary = -1 Then
                DtaSB.Fill(TblSB)
            Else
                DtaSB.FillBy_PrimarySubBrand(TblSB, vPrimary)
            End If
            If TblSB.Count > 0 Then
                If vClear = True Then
                    Me.DgvSubBrand.Rows.Clear()
                Else
                    For Each Me.RowDgv In Me.DgvSubBrand.Rows
                        ListPrimary.Add(Me.RowDgv.Cells(Me.ColPrimary_SB.Index).Value)
                    Next
                End If
                For Each RowSB In TblSB
                    If ListPrimary.Contains(RowSB.Primary) = False Then
                        Me.DgvSubBrand.Rows.Add()
                        RowDgv = Me.DgvSubBrand.Rows(Me.DgvSubBrand.Rows.Count - 1)
                        RowDgv.Cells(Me.ColNo_SB.Index).Value = ""
                        RowDgv.Cells(Me.ColCode_SB.Index).Value = RowSB.Group_Code
                        RowDgv.Cells(Me.ColDesc_SB.Index).Value = RowSB.Description
                        RowDgv.Cells(Me.ColPrint_SB.Index).Value = Boolean.FalseString
                        RowDgv.Cells(Me.ColPrimary_SB.Index).Value = RowSB.Primary
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub LoadAllCategory(ByVal vPrimary As Double, ByVal vClear As Boolean)
        Dim DtaC As New DseProductListReportTableAdapters.tblmaster_categoryTableAdapter
        Dim TblC As New DseProductListReport.tblmaster_categoryDataTable
        Dim RowC As DseProductListReport.tblmaster_categoryRow
        Dim ListPrimary As New List(Of Double)
        Try
            DtaC.Connection.ConnectionString = Core_Module.GetConstring
            If vPrimary = -1 Then
                DtaC.Fill(TblC)
            Else
                DtaC.FillBy_PrimarySubBrand(TblC, vPrimary)
            End If
            If TblC.Count > 0 Then
                If vClear = True Then
                    Me.DgvCategory.Rows.Clear()
                Else
                    For Each Me.RowDgv In Me.DgvCategory.Rows
                        ListPrimary.Add(Me.RowDgv.Cells(Me.ColPrimary_C.Index).Value)
                    Next
                End If
                For Each RowC In TblC
                    If ListPrimary.Contains(RowC.Primary) = False Then
                        Me.DgvCategory.Rows.Add()
                        RowDgv = Me.DgvCategory.Rows(Me.DgvCategory.Rows.Count - 1)
                        RowDgv.Cells(Me.ColNo_C.Index).Value = ""
                        RowDgv.Cells(Me.ColCode_C.Index).Value = RowC.Category_Code
                        RowDgv.Cells(Me.ColDesc_C.Index).Value = RowC.Description
                        RowDgv.Cells(Me.ColPrint_C.Index).Value = Boolean.FalseString
                        RowDgv.Cells(Me.ColPrimary_C.Index).Value = RowC.Primary
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub LoadAllProduct(ByVal vPrimary As Double, ByVal vClear As Boolean)
        Dim DtaP As New DseProductListReportTableAdapters.tblmaster_productTableAdapter
        Dim TblP As New DseProductListReport.tblmaster_productDataTable
        Dim RowP As DseProductListReport.tblmaster_productRow
        Dim ListPrimary As New List(Of Double)
        Try
            DtaP.Connection.ConnectionString = Core_Module.GetConstring
            If vPrimary = -1 Then
                DtaP.Fill(TblP)
            Else
                DtaP.FillBy_PrimaryCategory(TblP, vPrimary)
            End If
            If TblP.Count > 0 Then
                If vClear = True Then
                    Me.DgvProduct.Rows.Clear()
                Else
                    For Each Me.RowDgv In Me.DgvProduct.Rows
                        ListPrimary.Add(Me.RowDgv.Cells(Me.ColPrimary_P.Index).Value)
                    Next
                End If
                For Each RowP In TblP
                    If ListPrimary.Contains(RowP.Primary) = False Then
                        Me.DgvProduct.Rows.Add()
                        RowDgv = Me.DgvProduct.Rows(Me.DgvProduct.Rows.Count - 1)
                        RowDgv.Cells(Me.ColNo_P.Index).Value = ""
                        RowDgv.Cells(Me.ColCode_P.Index).Value = RowP.Product_Code
                        RowDgv.Cells(Me.ColDesc_P.Index).Value = RowP.Description
                        RowDgv.Cells(Me.ColPrint_P.Index).Value = Boolean.FalseString
                        RowDgv.Cells(Me.ColPrimary_C.Index).Value = RowP.Primary
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub
#End Region

#Region "check uncheck all"
    ''-------------PG::
    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Threading.Thread.Sleep(500)
            Me.Cursor = Cursors.WaitCursor
            Me.Check(Me.DgvProductGroup, Me.ColPrint_PG.Index, True)
            Me.DgvBrand.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvBrand.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each Me.RowDgv In Me.DgvProductGroup.Rows
                Me.Show_(Me.RowDgv.Cells(Me.ColPrimary_PG.Index).Value, EnumOption.PG, False)
            Next
            Me.DgvBrand.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvBrand.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.ObjCd.SetNumber(Me.DgvBrand)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            Threading.Thread.Sleep(500)
            Me.Check(Me.DgvProductGroup, Me.ColPrint_PG.Index, False)
            Me.DgvBrand.Rows.Clear()
            Me.DgvSubBrand.Rows.Clear()
            Me.DgvCategory.Rows.Clear()
            Me.DgvProduct.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub
    ''-------------B::
    Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Try
            Threading.Thread.Sleep(500)

            Me.Cursor = Cursors.WaitCursor
            Me.Check(Me.DgvBrand, Me.ColPrint_B.Index, True)
            Me.DgvSubBrand.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvSubBrand.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each Me.RowDgv In Me.DgvBrand.Rows
                Me.Show_(Me.RowDgv.Cells(Me.ColPrimary_B.Index).Value, EnumOption.B, False)
            Next
            Me.DgvSubBrand.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvSubBrand.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.ObjCd.SetNumber(Me.DgvSubBrand)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Try
            Threading.Thread.Sleep(500)

            Me.Check(Me.DgvBrand, Me.ColPrint_B.Index, False)
            Me.DgvSubBrand.Rows.Clear()
            Me.DgvCategory.Rows.Clear()
            Me.DgvProduct.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub
    ''-------------SB:
    Private Sub LinkLabel6_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel6.LinkClicked
        Try
            Threading.Thread.Sleep(500)

            Me.Cursor = Cursors.WaitCursor
            Me.Check(Me.DgvSubBrand, Me.ColPrint_SB.Index, True)
            Me.DgvCategory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvCategory.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each Me.RowDgv In Me.DgvSubBrand.Rows
                Me.Show_(Me.RowDgv.Cells(Me.ColPrimary_SB.Index).Value, EnumOption.SB, False)
            Next
            Me.DgvCategory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvCategory.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.ObjCd.SetNumber(Me.DgvCategory)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LinkLabel5_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel5.LinkClicked
        Try
            Threading.Thread.Sleep(500)

            Me.Check(Me.DgvSubBrand, Me.ColPrint_SB.Index, False)
            Me.DgvCategory.Rows.Clear()
            Me.DgvProduct.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub
    ''-------------C:
    Private Sub LinkLabel8_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel8.LinkClicked
        Try
            'Dim StartT As Date
            'StartT = Now
            Threading.Thread.Sleep(500)


            Me.Cursor = Cursors.WaitCursor
            Me.Check(Me.DgvCategory, Me.ColPrint_C.Index, True)
            Me.DgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvProduct.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each Me.RowDgv In Me.DgvCategory.Rows
                Me.Show_(Me.RowDgv.Cells(Me.ColPrimary_C.Index).Value, EnumOption.C, False)
            Next
            Me.DgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvProduct.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.ObjCd.SetNumber(Me.DgvProduct)


            'Dim TS As TimeSpan = Now - StartT
            'If TS.TotalSeconds > 1 Then
            '    MsgBox(ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds")
            'Else
            '    MsgBox(ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds")
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LinkLabel7_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel7.LinkClicked
        Try
            Threading.Thread.Sleep(500)

            Me.Check(Me.DgvCategory, Me.ColPrint_C.Index, False)
            Me.DgvProduct.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    ''-------------common
    Private Sub Check(ByVal vDgv As DataGridView, ByVal vCol As Integer, ByVal vCheck As Boolean)
        Try
            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each Me.RowDgv In vDgv.Rows
                If vCheck = True Then
                    Me.RowDgv.Cells(vCol).Value = Boolean.TrueString
                Else
                    Me.RowDgv.Cells(vCol).Value = Boolean.FalseString
                End If
            Next
            vDgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub
#End Region

#Region "getlist"
    Private Sub GetListBrand(ByVal vPrimaryPG As Double)
        Dim DtaB As New DseProductListReportTableAdapters.tblmaster_brandTableAdapter
        Dim TblB As New DseProductListReport.tblmaster_brandDataTable
        Dim RowB As DseProductListReport.tblmaster_brandRow
        Try
            DtaB.Connection.ConnectionString = Core_Module.GetConstring
            DtaB.FillBy_PrimaryPG(TblB, vPrimaryPG)
            For Each RowB In TblB
                ListBrand.Add(RowB.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub GetListSubBrand(ByVal vListBrand As List(Of Double))
        Dim DtaSB As New DseProductListReportTableAdapters.tblmaster_subbrandTableAdapter
        Dim TblSB As New DseProductListReport.tblmaster_subbrandDataTable
        Dim RowSB As DseProductListReport.tblmaster_subbrandRow
        Try
            DtaSB.Connection.ConnectionString = Core_Module.GetConstring
            For Each d As Double In vListBrand
                DtaSB.FillBy_PrimarySubBrand(TblSB, d)
                For Each RowSB In TblSB
                    ListSubBrand.Add(RowSB.Primary)
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub GetListCategory(ByVal vListSubBrand As List(Of Double))
        Dim DtaC As New DseProductListReportTableAdapters.tblmaster_categoryTableAdapter
        Dim TblC As New DseProductListReport.tblmaster_categoryDataTable
        Dim RowC As DseProductListReport.tblmaster_categoryRow
        Try
            DtaC.Connection.ConnectionString = Core_Module.GetConstring
            For Each d As Double In vListSubBrand
                DtaC.FillBy_PrimarySubBrand(TblC, d)
                For Each RowC In TblC
                    ListCategory.Add(RowC.Primary)
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub GetListProduct(ByVal vListCategory As List(Of Double))
        Dim DtaP As New DseProductListReportTableAdapters.tblmaster_productTableAdapter
        Dim TblP As New DseProductListReport.tblmaster_productDataTable
        Dim RowP As DseProductListReport.tblmaster_productRow
        Try
            DtaP.Connection.ConnectionString = Core_Module.GetConstring
            For Each d As Double In vListCategory
                DtaP.FillBy_PrimaryCategory(TblP, d)
                For Each RowP In TblP
                    ListProduct.Add(RowP.Primary)
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub
#End Region

#Region "dgv edit"
    Private Sub DgvProductGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvProductGroup.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColPrint_PG.Index Then
                e.Cancel = True
            Else
                Me.Buffer = Me.DgvProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub DgvProductGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvProductGroup.CellEndEdit
        Try
            If Me.Buffer <> Me.DgvProductGroup.Item(Me.ColPrint_PG.Index, e.RowIndex).Value Then
                Me.DgvBrand.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvBrand.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                If Me.DgvProductGroup.Item(Me.ColPrint_PG.Index, e.RowIndex).Value = Boolean.TrueString Then
                    Me.Show_(Me.DgvProductGroup.Item(Me.ColPrimary_PG.Index, e.RowIndex).Value, EnumOption.PG, False)
                Else
                    Me.UnShow(Me.DgvProductGroup.Item(Me.ColPrimary_PG.Index, e.RowIndex).Value, EnumOption.PG)
                End If
                Me.DgvBrand.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvBrand.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.ObjCd.SetNumber(Me.DgvBrand)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub DgvBrand_CellBeginEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvBrand.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColPrint_B.Index Then
                e.Cancel = True
            Else
                Me.Buffer = Me.DgvBrand.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub DgvBrand_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvBrand.CellEndEdit
        Try
            If Me.Buffer <> Me.DgvBrand.Item(Me.ColPrint_PG.Index, e.RowIndex).Value Then
                Me.DgvSubBrand.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvSubBrand.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                If Me.DgvBrand.Item(Me.ColPrint_B.Index, e.RowIndex).Value = Boolean.TrueString Then
                    Me.Show_(Me.DgvBrand.Item(Me.ColPrimary_B.Index, e.RowIndex).Value, EnumOption.B, False)
                Else
                    Me.UnShow(Me.DgvBrand.Item(Me.ColPrimary_B.Index, e.RowIndex).Value, EnumOption.B)
                End If
                Me.DgvSubBrand.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvSubBrand.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.ObjCd.SetNumber(Me.DgvSubBrand)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub DgvSubBrand_CellBeginEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvSubBrand.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColPrint_SB.Index Then
                e.Cancel = True
            Else
                Me.Buffer = Me.DgvSubBrand.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub DgvSubBrand_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvSubBrand.CellEndEdit
        Try
            If Me.Buffer <> Me.DgvSubBrand.Item(Me.ColPrint_PG.Index, e.RowIndex).Value Then
                Me.DgvCategory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvCategory.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                If Me.DgvSubBrand.Item(Me.ColPrint_SB.Index, e.RowIndex).Value = Boolean.TrueString Then
                    Me.Show_(Me.DgvSubBrand.Item(Me.ColPrimary_SB.Index, e.RowIndex).Value, EnumOption.SB, False)
                Else
                    Me.UnShow(Me.DgvSubBrand.Item(Me.ColPrimary_SB.Index, e.RowIndex).Value, EnumOption.SB)
                End If
                Me.DgvCategory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvCategory.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.ObjCd.SetNumber(Me.DgvCategory)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub DgvCategory_CellBeginEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvCategory.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColPrint_C.Index Then
                e.Cancel = True
            Else
                Me.Buffer = Me.DgvCategory.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub DgvCategory_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvCategory.CellEndEdit
        Try
            If Me.Buffer <> Me.DgvCategory.Item(Me.ColPrint_PG.Index, e.RowIndex).Value Then
                Me.DgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvProduct.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                If Me.DgvCategory.Item(Me.ColPrint_C.Index, e.RowIndex).Value = Boolean.TrueString Then
                    Me.Show_(Me.DgvCategory.Item(Me.ColPrimary_C.Index, e.RowIndex).Value, EnumOption.C, False)
                Else
                    Me.UnShow(Me.DgvCategory.Item(Me.ColPrimary_C.Index, e.RowIndex).Value, EnumOption.C)
                End If
                Me.DgvProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvProduct.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.ObjCd.SetNumber(Me.DgvProduct)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub
#End Region

    Private Sub BtnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint.Click
        Dim DS As New DseProductListReport
        Dim TblCetak As New DseProductListReport.TblCetakDataTable
        Dim TblCetak2 As DseProductListReport.TblCetakDataTable = DS.TblCetak
        Dim RowCetak As DseProductListReport.TblCetakRow
        Dim Report As New CrpProductListReport
        Dim Title As String = ""
        Dim FrmC As New FrmMasterProductListReport_Print
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.RboProductGroup.Checked = True Then
                For Each Me.RowDgv In Me.DgvProductGroup.Rows
                    RowCetak = TblCetak.NewTblCetakRow
                    RowCetak.Code = Me.RowDgv.Cells(Me.ColCode_PG.Index).Value
                    RowCetak.Desc = Me.RowDgv.Cells(Me.ColDesc_PG.Index).Value
                    TblCetak.AddTblCetakRow(RowCetak)
                Next
                Title = "Product Group List Report"
            ElseIf Me.RboBrand.Checked = True Then
                For Each Me.RowDgv In Me.DgvBrand.Rows
                    RowCetak = TblCetak.NewTblCetakRow
                    RowCetak.Code = Me.RowDgv.Cells(Me.ColCode_B.Index).Value
                    RowCetak.Desc = Me.RowDgv.Cells(Me.ColDesc_B.Index).Value
                    TblCetak.AddTblCetakRow(RowCetak)
                Next
                Title = "Brand List Report"
            ElseIf Me.RboSubBrand.Checked = True Then
                For Each Me.RowDgv In Me.DgvSubBrand.Rows
                    RowCetak = TblCetak.NewTblCetakRow
                    RowCetak.Code = Me.RowDgv.Cells(Me.ColCode_SB.Index).Value
                    RowCetak.Desc = Me.RowDgv.Cells(Me.ColDesc_SB.Index).Value
                    TblCetak.AddTblCetakRow(RowCetak)
                Next
                Title = "Sub Brand List Report"
            ElseIf Me.RboCategory.Checked = True Then
                For Each Me.RowDgv In Me.DgvCategory.Rows
                    RowCetak = TblCetak.NewTblCetakRow
                    RowCetak.Code = Me.RowDgv.Cells(Me.ColCode_C.Index).Value
                    RowCetak.Desc = Me.RowDgv.Cells(Me.ColDesc_C.Index).Value
                    TblCetak.AddTblCetakRow(RowCetak)
                Next
                Title = "Category List Report"
            ElseIf Me.RboProduct.Checked = True Then
                For Each Me.RowDgv In Me.DgvProduct.Rows
                    RowCetak = TblCetak.NewTblCetakRow
                    RowCetak.Code = Me.RowDgv.Cells(Me.ColCode_P.Index).Value
                    RowCetak.Desc = Me.RowDgv.Cells(Me.ColDesc_P.Index).Value
                    TblCetak.AddTblCetakRow(RowCetak)
                Next
                Title = "Product List Report"
            End If
            If TblCetak.Rows.Count > 0 Then
                TblCetak.DefaultView.Sort = "[Code] ASC"
                TblCetak2.Merge(TblCetak.DefaultView.ToTable)
                Report.SetDataSource(DS)
                Report.SetParameterValue("Title", Title)
                FrmC.CrystalReportViewer1.ReportSource = Report
                FrmC.CrystalReportViewer1.Zoom(90)
                FrmC.CrystalReportViewer1.Refresh()
                FrmC.CrystalReportViewer1.Show()
                FrmC.Text = Title
                FrmC.ShowDialog()
            Else
                Me.Cursor = Cursors.Default
                MsgBox("Load data first", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmC.Dispose()
            Me.Cursor = Cursors.Default
        End Try
    End Sub

#Region "dgv sorted"

    Private Sub DgvProductGroup_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvProductGroup.Sorted
        Try
            Me.ObjCd.SetNumber(Me.DgvProductGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub DgvBrand_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvBrand.Sorted
        Try
            Me.ObjCd.SetNumber(Me.DgvBrand)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub DgvSubBrand_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvSubBrand.Sorted
        Try
            Me.ObjCd.SetNumber(Me.DgvSubBrand)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub DgvCategory_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvCategory.Sorted
        Try
            Me.ObjCd.SetNumber(Me.DgvCategory)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub DgvProduct_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvProduct.Sorted
        Try
            Me.ObjCd.SetNumber(Me.DgvProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub
#End Region

End Class
