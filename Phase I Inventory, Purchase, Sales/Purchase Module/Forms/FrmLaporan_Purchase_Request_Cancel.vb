Option Explicit On
Option Strict Off
Public Class FrmLaporan_Purchase_Request_Cancel

    Public BAckColorAll As Color = Core_Module.GetSMESSBackColorAll
    Private ObjCd As New Core_Module.ClsCekData
    Private ObjRefParty As New Core_Module.ClsMasterRefParty
    Private ObjPR As New ClsTransaksiPR
    Private ObjUser As New Core_Module.ClsUser
    Private ObjCancel As New ClsTransaksiPR_Cancel
    Private ObjPickupOrder As New ClsTransaksiPR_PickUpOrder
    Dim ObjLicense As New Core_Module.ClsLicense

    '''Product Code
    Private Sub TxtProductCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtProductCode.KeyDown
        Dim ObjProduct As New Core_Module.clsMasterProduct

        Try
            If e.KeyValue = 13 Then
                If Me.TxtProductCode.Text <> "" Then
                    If ObjProduct.ApakahSudahAda(Me.TxtProductCode.Text) = True Then
                        Me.TxtPrimaryProduct.Text = ObjProduct.GetPrimary(Me.TxtProductCode.Text)
                        Me.TxtProductCode.Text = ObjProduct.GetProductCode(Me.TxtPrimaryProduct.Text)
                        Me.TxtDescProduct.Text = ObjProduct.GetDescription(Me.TxtPrimaryProduct.Text)
                        Me.TxtRequestBy.Focus()
                    Else
                        Me.btnFindProduct.PerformClick()
                    End If
                Else
                    Me.btnFindProduct.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtProductCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtProductCode.Leave
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim hj As New Core_Module.frmCariMasterProduct
        hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.TxtProductCode.Text <> "" Then
                If ObjProduct.ApakahSudahAda(Me.TxtProductCode.Text) = True Then
                    Me.TxtPrimaryProduct.Text = ObjProduct.GetPrimary(Me.TxtProductCode.Text)
                    Me.TxtDescProduct.Text = ObjProduct.GetDescription(Me.TxtPrimaryProduct.Text)
                ElseIf ObjProduct.ApakahSudahAda(Me.TxtProductCode.Text) = False Then
                    If Me.TxtProductCode.Text <> "" Then
                        hj.txtCari.Text = Me.TxtProductCode.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryProduct > 0 Then
                        Me.TxtPrimaryProduct.Text = hj.PrimaryProduct
                        Me.TxtProductCode.Text = ObjProduct.GetProductCode(Me.TxtPrimaryProduct.Text)
                        Me.TxtDescProduct.Text = ObjProduct.GetDescription(Me.TxtPrimaryProduct.Text)
                    Else
                        Me.TxtPrimaryProduct.Text = "-1"
                        Me.TxtProductCode.Clear()
                        Me.TxtDescProduct.Clear()
                        Me.TxtProductCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryProduct.Text = "-1"
                Me.TxtProductCode.Clear()
                Me.TxtDescProduct.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindProduct.Click
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim hj As New Core_Module.frmCariMasterProduct
        Try
            hj.BackColorAll = Me.BAckColorAll
            If Me.TxtProductCode.Text <> "" Then
                hj.txtCari.Text = Me.TxtProductCode.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryProduct > 0 Then
                Me.TxtPrimaryProduct.Text = hj.PrimaryProduct
                Me.TxtProductCode.Text = ObjProduct.GetProductCode(Me.TxtPrimaryProduct.Text)
                Me.TxtDescProduct.Text = ObjProduct.GetDescription(Me.TxtPrimaryProduct.Text)
                Me.TxtRequestBy.Focus()
            Else
                Me.TxtPrimaryProduct.Text = "-1"
                Me.TxtProductCode.Clear()
                Me.TxtDescProduct.Clear()
                Me.TxtProductCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    '''RefParty
    Private Sub TxtRequestBy_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRequestBy.KeyDown
        Dim ObjRefParty As New Core_Module.ClsMasterRefParty
        Try
            If e.KeyValue = 13 Then
                If Me.TxtRequestBy.Text <> "" Then
                    If ObjRefParty.ApakahSudahAda(Me.TxtRequestBy.Text) = True Then
                        Me.TxtPrimaryRefParty.Text = ObjRefParty.GetPrimary(Me.TxtRequestBy.Text)
                        Me.TxtDescRefParty.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                        Me.txtNoPR.Focus()
                    Else
                        Me.BtnFindRefParty.PerformClick()
                    End If
                Else
                    Me.BtnFindRefParty.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRequestBy_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRequestBy.Leave
        Dim ObjRefParty As New Core_Module.ClsMasterRefParty
        Dim hj As New Core_Module.FrmMasterRefParty_Cari
        hj.Aktif = True
        hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.TxtRequestBy.Text <> "" Then
                If ObjRefParty.ApakahSudahAda(Me.TxtRequestBy.Text) = True Then
                    Me.TxtPrimaryRefParty.Text = ObjRefParty.GetPrimary(Me.TxtRequestBy.Text)
                    Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                    Me.TxtDescRefParty.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                    Me.txtNoPR.Focus()
                ElseIf ObjRefParty.ApakahSudahAda(Me.TxtRequestBy.Text) = False Then
                    If Me.TxtRequestBy.Text <> "" Then
                        hj.txtCari.Text = Me.TxtRequestBy.Text
                    End If

                    hj.ShowDialog()
                    If hj.PrimaryRefParty > 0 Then
                        Me.TxtPrimaryRefParty.Text = hj.PrimaryRefParty
                        Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                        Me.TxtDescRefParty.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                        Me.txtNoPR.Focus()
                    Else
                        Me.TxtPrimaryRefParty.Text = "-1"
                        Me.TxtRequestBy.Clear()
                        Me.TxtDescRefParty.Clear()
                        Me.TxtRequestBy.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryRefParty.Text = "-1"
                Me.TxtRequestBy.Clear()
                Me.TxtDescRefParty.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindRefParty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindRefParty.Click
        Dim ObjRefParty As New Core_Module.ClsMasterRefParty
        Dim hj As New Core_Module.FrmMasterRefParty_Cari
        hj.Aktif = True
        Try
            hj.BackColorAll = Me.BAckColorAll
            If Me.TxtRequestBy.Text <> "" Then
                hj.txtCari.Text = Me.TxtRequestBy.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryRefParty > 0 Then
                Me.TxtPrimaryRefParty.Text = hj.PrimaryRefParty
                Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                Me.TxtDescRefParty.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                Me.txtNoPR.Focus()
            Else
                Me.TxtPrimaryRefParty.Text = "-1"
                Me.TxtRequestBy.Clear()
                Me.TxtDescRefParty.Clear()
                Me.TxtRequestBy.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    '''PR
    Private Sub txtNoPR_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNoPR.KeyDown
        Dim objPR As New ClsTransaksiPR
        Dim jk As New ClsTransaksiPR_Item
        Try
            If Me.txtNoPR.ReadOnly = True Then Exit Sub
            If e.KeyValue = 13 Then
                If Me.txtNoPR.Text <> "" Then
                    If objPR.ApakahSudahAda(Me.txtNoPR.Text) = True Then
                        Me.TxtPrimaryPR.Text = objPR.GetPrimary(Me.txtNoPR.Text)
                        Me.txtNoPR.Text = objPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                        'Me.TxtPrimaryRefParty.Text = objPR.GetPrimaryRefParty(Me.TxtPrimaryPR.Text)
                        'Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                        'Me.TxtDescRefParty.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                        Me.btnFilter.Focus()
                    Else
                        Me.BtnFindPR.PerformClick()
                    End If
                Else
                    Me.BtnFindPR.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoPR_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoPR.Leave
        Dim ObjPR As New ClsTransaksiPR
        Dim hj As New FrmTransaksiPR_Cari
        Dim jk As New ClsTransaksiPR_Item
        hj.vRFP = False
        hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.txtNoPR.ReadOnly = True Then Exit Sub

            If Me.txtNoPR.Text <> "" Then
                If ObjPR.ApakahSudahAda(Me.txtNoPR.Text) = True Then
                    Me.TxtPrimaryPR.Text = ObjPR.GetPrimary(Me.txtNoPR.Text)
                    Me.txtNoPR.Text = ObjPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                    'Me.TxtPrimaryRefParty.Text = ObjPR.GetPrimaryRefParty(Me.TxtPrimaryPR.Text)
                    'Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                    'Me.TxtDescRefParty.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                    Me.btnFilter.Focus()
                Else
                    If Me.txtNoPR.Text <> "" Then
                        hj.txtCari.Text = Me.txtNoPR.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryPR > 0 Then
                        Me.TxtPrimaryPR.Text = hj.PrimaryPR
                        Me.txtNoPR.Text = ObjPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                        'Me.TxtPrimaryRefParty.Text = ObjPR.GetPrimaryRefParty(Me.TxtPrimaryPR.Text)
                        'Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                        'Me.TxtDescRefParty.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                        Me.btnFilter.Focus()
                    Else
                        Me.TxtPrimaryPR.Text = "-1"
                        Me.txtNoPR.Clear()
                        'Me.TxtPrimaryRefParty.Text = "-1"
                        'Me.TxtRequestBy.Clear()
                        'Me.TxtDescRefParty.Clear()
                    End If
                End If
            Else
                Me.TxtPrimaryPR.Text = "-1"
                Me.txtNoPR.Clear()
                'Me.TxtPrimaryRefParty.Text = "-1"
                'Me.TxtRequestBy.Clear()
                'Me.TxtDescRefParty.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindPR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindPR.Click

        Dim ObjPR As New ClsTransaksiPR
        Dim hj As New FrmTransaksiPR_Cari
        Dim jk As New ClsTransaksiPR_Item
        hj.vRFP = False
        Try

            hj.BackColorAll = Me.BAckColorAll
            If Me.txtNoPR.Text <> "" Then
                hj.txtCari.Text = Me.txtNoPR.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryPR > 0 Then
                Me.TxtPrimaryPR.Text = hj.PrimaryPR
                Me.txtNoPR.Text = ObjPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                'Me.TxtPrimaryRefParty.Text = ObjPR.GetPrimaryRefParty(Me.TxtPrimaryPR.Text)
                'Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                'Me.TxtDescRefParty.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
            Else
                Me.TxtPrimaryPR.Text = "-1"
                Me.txtNoPR.Clear()
                Me.txtNoPR.Focus()
                'Me.TxtPrimaryRefParty.Text = "-1"
                'Me.TxtRequestBy.Clear()
                'Me.TxtDescRefParty.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    '''Got Lost Focus
    Private Sub TxtRequestBy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRequestBy.GotFocus, _
    TxtProductCode.GotFocus, txtNoPR.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Txt.Enabled = True Then
                    Txt.BackColor = Me.BAckColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRequestBy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRequestBy.LostFocus, _
    txtNoPR.LostFocus, TxtProductCode.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Txt.Enabled = True Then
                    Txt.BackColor = Me.TxtPrimaryProduct.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.DtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.DtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.DtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.DtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.DtpFrom.Value.Date > Me.DtpTo.Value.Date Then
                MsgBox("Starting period must be smaller than ending period", MsgBoxStyle.Information, "Information")
                Me.DtpFrom.Focus()
                Exit Sub
            End If
            Me.Cetakan(Me.DtpFrom.Value, Me.DtpTo.Value, Me.TxtPrimaryRefParty.Text, Me.TxtPrimaryProduct.Text, Me.TxtPrimaryPR.Text)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub FrmLaporan_Purchase_Request_Cancel_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Me.WindowState = FormWindowState.Maximized
        Me.DtpFrom.Select()
        Me.DtpFrom.Value = Now
        Me.DtpTo.Value = Now
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub


    Public Sub Cetakan(ByVal StartDate As Date, ByVal EndDate As Date, ByVal PrimaryRefParty As Double, ByVal PrimaryProduct As Double, ByVal PrimaryPR As Double)
        Dim ReportDetail As New CrpPurchase_Request_Cancel_Report
        Dim DS As New DsePurchaseRequest_CancelReport

        Dim DTACetak_Header As New DsePurchaseRequest_CancelReportTableAdapters.TblCetak_HeaderTableAdapter
        Dim TBLCetak_Header As DsePurchaseRequest_CancelReport.TblCetak_HeaderDataTable = DS.TblCetak_Header
        Dim RowCetak_Header As DsePurchaseRequest_CancelReport.TblCetak_HeaderRow

        Dim DTACetak_Item As New DsePurchaseRequest_CancelReportTableAdapters.TblCetak_ItemTableAdapter
        Dim TBLCetak_Item As DsePurchaseRequest_CancelReport.TblCetak_ItemDataTable = DS.TblCetak_Item
        Dim RowCetak_Item As DsePurchaseRequest_CancelReport.TblCetak_ItemRow

        Dim TBLHeader As DsePurchaseRequest_CancelReport.TblHeaderDataTable = DS.TblHeader
        Dim TBLItem As DsePurchaseRequest_CancelReport.TblItemDataTable = DS.TblItem

        Dim ROWHeader As DsePurchaseRequest_CancelReport.TblHeaderRow
        Dim ROWItem As DsePurchaseRequest_CancelReport.TblItemRow

        Dim paramStartDate As Date
        Dim paramEndDate As Date
        Dim PrimaryMax As Double = -1
        Dim PrimaryMin As Double = -1
        Dim PrimaryTransaksi As Double = -1
        Dim index As Integer = 0
        Dim Filter1 As String = ""
        Dim Filter2 As String = ""
        Dim DV1 As DataRowView
        Dim DV2 As DataRowView
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            TBLHeader.Rows.Clear()
            TBLItem.Rows.Clear()
            DTACetak_Header.Connection.ConnectionString = Core_Module.GetConstring
            DTACetak_Item.Connection.ConnectionString = DTACetak_Header.Connection.ConnectionString

            paramStartDate = Me.DtpFrom.Value
            paramEndDate = Me.DtpTo.Value
            DTACetak_Header.FillBy_StartEndDate(TBLCetak_Header, paramStartDate, paramEndDate)

            TBLCetak_Header.DefaultView.Sort = Nothing
            TBLCetak_Header.DefaultView.Sort = "[Primary]"
            If TBLCetak_Header.DefaultView.Count > 0 Then
                RowCetak_Header = TBLCetak_Header.DefaultView(0).Row
                PrimaryMin = RowCetak_Header.primary
                RowCetak_Header = TBLCetak_Header.DefaultView(TBLCetak_Header.DefaultView.Count - 1).Row
                PrimaryMax = RowCetak_Header.primary
            End If


            DTACetak_Item.FillBy_PrimaryMinAndMax(TBLCetak_Item, PrimaryMin, PrimaryMax)
            Filter1 = "[Primary] = [Primary]"
            If PrimaryPR <> -1 Then
                Filter1 &= " AND [primary_pr] = " & PrimaryPR & ""
            End If

            If PrimaryRefParty <> -1 Then
                Filter1 &= " AND [primary_refparty] = " & PrimaryRefParty & ""
            End If

            TBLCetak_Header.DefaultView.Sort = Nothing
            TBLCetak_Header.DefaultView.Sort = "[tanggal], document_no"
            TBLCetak_Header.DefaultView.RowFilter = Nothing
            TBLCetak_Header.DefaultView.RowFilter = Filter1
            If TBLCetak_Header.DefaultView.Count > 0 Then
                For Each DV1 In TBLCetak_Header.DefaultView
                    RowCetak_Header = DV1.Row
                    ROWHeader = TBLHeader.NewTblHeaderRow
                    '''isi header
                    With RowCetak_Header
                        ROWHeader.Primary = .primary
                        ROWHeader.Document_No = .document_no
                        ROWHeader.PR_No = .pr_no
                        ROWHeader._Date = Format(.tanggal, "dd MMM yyyy")
                        ROWHeader.Reason = .reason
                        ROWHeader.Request_By = .request_by & " - " & .desc_ref_party
                        ROWHeader.Remarks = .remarks
                    End With
                    TBLHeader.AddTblHeaderRow(ROWHeader)

                    ''filter buat itemnya
                    If PrimaryProduct <> -1 Then
                        Filter2 = "[primary transaksi] = " & RowCetak_Header.primary.ToString & " AND [primary code no] = " & PrimaryProduct & ""
                    Else
                        Filter2 = "[primary transaksi] = " & RowCetak_Header.primary.ToString
                    End If

                    TBLCetak_Item.DefaultView.RowFilter = Nothing
                    TBLCetak_Item.DefaultView.RowFilter = Filter2
                    If TBLCetak_Item.DefaultView.Count > 0 Then
                        For Each DV2 In TBLCetak_Item.DefaultView
                            RowCetak_Item = DV2.Row
                            '''isi item
                            With RowCetak_Item
                                ROWItem = TBLItem.NewTblItemRow
                                ROWItem.Primary = .primary
                                ROWItem.Primary_Transaksi = RowCetak_Header.primary
                                ROWItem.Primary_Product = .primary_code_no
                                ROWItem.Product_Code = .Product_Code
                                ROWItem.Description = .Description
                                ROWItem._Decimal = ._Decimal
                                ROWItem.Qty = ObjCd.DigitGrouping(.qty, ROWItem._Decimal)
                                ROWItem.Unit_Code = .Unit_Code
                                ROWItem.Remarks = .remarks
                                TBLItem.AddTblItemRow(ROWItem)
                            End With
                        Next
                    End If
                Next
            End If
            ReportDetail.SetDataSource(DS)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_StartDate.ParameterFieldName, paramStartDate)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_EndDate.ParameterFieldName, paramEndDate)
            Me.crv.ReportSource = ReportDetail
            Me.crv.Show()
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
    End Sub
End Class