Option Explicit On
Option Strict Off
Public Class FrmLaporan_PR_Balance

    Public BAckColorAll As Color = Core_Module.GetSMESSBackColorAll
    Private ObjCd As New Core_Module.ClsCekData
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
                        Me.btnFilter.Focus()
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
                    Me.btnFilter.Focus()
                ElseIf ObjRefParty.ApakahSudahAda(Me.TxtRequestBy.Text) = False Then
                    If Me.TxtRequestBy.Text <> "" Then
                        hj.txtCari.Text = Me.TxtRequestBy.Text
                    End If

                    hj.ShowDialog()
                    If hj.PrimaryRefParty > 0 Then
                        Me.TxtPrimaryRefParty.Text = hj.PrimaryRefParty
                        Me.TxtRequestBy.Text = ObjRefParty.GetCode(Me.TxtPrimaryRefParty.Text)
                        Me.TxtDescRefParty.Text = ObjRefParty.GetDesc(Me.TxtPrimaryRefParty.Text)
                        Me.btnFilter.Focus()
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
                Me.btnFilter.Focus()
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

    '''Got Lost Focus
    Private Sub TxtRequestBy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRequestBy.GotFocus, TxtProductCode.GotFocus
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

    Private Sub TxtRequestBy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRequestBy.LostFocus, TxtProductCode.LostFocus
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

    Private Sub FrmLaporan_PR_Balance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.WindowState = FormWindowState.Maximized
        Me.RboProduct.Checked = True
        Me.DtpFrom.Enabled = False
        Me.DtpTo.Enabled = False
        Me.cboCodeNo.SelectedIndex = 0
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
            Else
                Me.Cetakan(Me.DtpFrom.Value, Me.DtpTo.Value, Me.TxtPrimaryRefParty.Text, Me.TxtPrimaryProduct.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub ChkPeriod_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkPeriod.Click
        Try
            If Me.ChkPeriod.Checked = True Then
                Me.DtpFrom.Enabled = True
                Me.DtpTo.Enabled = True
            Else
                Me.DtpFrom.Enabled = False
                Me.DtpTo.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Cetakan(ByVal StartDate As Date, ByVal EndDate As Date, ByVal PrimaryRefParty As Double, ByVal PrimaryProduct As Double)
        Dim ReportDetail As New CrpPurchase_Request_Balance
        Dim DS As New DsePR_BalanceReport

        'Dim DTACetak_Header As New DsePR_BalanceReportTableAdapters.TblCetak_HeaderTableAdapter
        'Dim TBLCetak_Header As DsePR_BalanceReport.TblCetak_HeaderDataTable = DS.TblCetak_Header
        'Dim RowCetak_Header As DsePR_BalanceReport.TblCetak_HeaderRow

        Dim DTACetak_Item As New DsePR_BalanceReportTableAdapters.TblCetak_ItemTableAdapter
        Dim TBLCetak_Item As DsePR_BalanceReport.TblCetak_ItemDataTable = DS.TblCetak_Item
        Dim RowCetak_Item As DsePR_BalanceReport.TblCetak_ItemRow

        'Dim TBLHeader As DsePR_BalanceReport.TBLHeaderDataTable = DS.TBLHeader
        Dim TBLItem As DsePR_BalanceReport.TBLItemDataTable = DS.TBLItem

        'Dim ROWHeader As DsePR_BalanceReport.TBLHeaderRow
        Dim ROWItem As DsePR_BalanceReport.TBLItemRow

        Dim FilterByDate As Boolean = False
        Dim paramStartDate As Date
        Dim paramEndDate As Date
        Dim ParamJudulGroup As String = ""
        Dim ParamHeaderGroupBy As String = ""
        'Dim PrimaryMax As Double = -1
        'Dim PrimaryMin As Double = -1
        Dim PrimaryTransaksi As Double = -1
        Dim index As Integer = 0
        'Dim Filter1 As String = ""
        Dim Filter2 As String = ""
        'Dim DV1 As DataRowView
        Dim DV2 As DataRowView

        Try
            Windows.Forms.Cursor.Current = Cursors.Arrow

            'TBLHeader.Rows.Clear()
            TBLItem.Rows.Clear()
            'DTACetak_Header.Connection.ConnectionString = Core_Module.GetConstring
            DTACetak_Item.Connection.ConnectionString = Core_Module.GetConstring

            paramStartDate = Me.DtpFrom.Value
            paramEndDate = Me.DtpTo.Value

            '''untuk membedakan judul yang di group berdasarkan apa 
            If Me.RboProduct.Checked = True And Me.RboRequestBy.Checked = False Then
                ParamJudulGroup = "By Product"
                ParamHeaderGroupBy = "Ref. Party"
            ElseIf Me.RboRequestBy.Checked = True And Me.RboProduct.Checked = False Then
                ParamJudulGroup = "By Ref Party"
                ParamHeaderGroupBy = "Description"
            End If

            If Me.ChkPeriod.Checked = True Then
                '''klo pilih period maka fill sesua period
                'DTACetak_Header.FillBy_StartEndDate(TBLCetak_Header, paramStartDate, paramEndDate)
                FilterByDate = True
                DTACetak_Item.FillBy_TglSampaiDari(TBLCetak_Item, Me.DtpTo.Value, Me.DtpFrom.Value)
            Else
                '''klo gg pilih berarti fill semua 
                'DTACetak_Header.Fill(TBLCetak_Header)
                FilterByDate = False
                DTACetak_Item.Fill(TBLCetak_Item)
            End If

            Filter2 = ""
            If Me.TxtProductCode.Text.Trim <> "" Then
                If Me.TxtPrimaryRefParty.Text <> -1 Then
                    Filter2 = "[primary code no] = " + Me.TxtPrimaryProduct.Text + " AND [primary ref party] = " + Me.TxtPrimaryRefParty.Text
                Else
                    Filter2 = "[primary code no] = " + Me.TxtPrimaryProduct.Text
                End If
            Else
                If Me.TxtPrimaryRefParty.Text <> -1 Then
                    Filter2 = "[primary ref party] = " + Me.TxtPrimaryRefParty.Text
                End If
            End If

            TBLCetak_Item.DefaultView.RowFilter = Nothing
            TBLCetak_Item.DefaultView.RowFilter = Filter2
            If TBLCetak_Item.DefaultView.Count > 0 Then
                For Each DV2 In TBLCetak_Item.DefaultView
                    RowCetak_Item = DV2.Row
                    '''isi item
                    With RowCetak_Item
                        ROWItem = TBLItem.NewTBLItemRow
                        ROWItem.Primary_Transaksi = .primary_transaksi
                        ROWItem.Primary_Product = .primary_code_no
                        ROWItem._Date = .tanggal
                        ROWItem.PR_No = .no_transaksi
                        ROWItem.Product_Code = .Product_Code
                        ROWItem.Description = .Description
                        ROWItem._Decimal = ._Decimal
                        ROWItem.Ordered = ObjCd.DigitGrouping(.ordered, ROWItem._Decimal)
                        ROWItem.PickUp = ObjCd.DigitGrouping(.qty_pick_up, ROWItem._Decimal)
                        ROWItem.Cancel = ObjCd.DigitGrouping(.qty_cancel, ROWItem._Decimal)
                        ROWItem.OB = ObjCd.DigitGrouping(.OB, ROWItem._Decimal)
                        ROWItem.Unit_Code = .Unit_Code
                        If FilterByDate = False Then
                            ROWItem.Filter_By_Date = ""
                        Else
                            ROWItem.Filter_By_Date = "ada"
                        End If
                        If .status = ClsTransaksiPR.EnumStatus.PR_Pending Then
                            ROWItem.Status = "PR Pending"
                        ElseIf .status = ClsTransaksiPR.EnumStatus.Purchase_On_Progress Then
                            ROWItem.Status = "On Progress"
                        Else
                            ROWItem.Status = "Closed"
                        End If
                        ROWItem.Remarks = .remarks

                        If ParamJudulGroup = "By Product" Then
                            If Me.cboCodeNo.SelectedIndex = 0 Then
                                ROWItem.Group_By = .Product_Code
                            ElseIf Me.cboCodeNo.SelectedIndex = 1 Then
                                ROWItem.Group_By = .Description
                            ElseIf Me.cboCodeNo.SelectedIndex = 2 Then
                                ROWItem.Group_By = .Product_Code & " - " & .Description
                            End If
                            ROWItem.Item_Group_By = .Request_By_Desc
                        Else
                            ROWItem.Group_By = .Request_By_Code & " - " & .Request_By_Desc
                            If Me.cboCodeNo.SelectedIndex = 0 Then
                                ROWItem.Item_Group_By = .Product_Code
                            ElseIf Me.cboCodeNo.SelectedIndex = 1 Then
                                ROWItem.Item_Group_By = .Description
                            Else
                                ROWItem.Item_Group_By = .Product_Code & " - " & .Description
                            End If
                        End If
                        TBLItem.AddTBLItemRow(ROWItem)
                    End With
                Next
            End If

            ReportDetail.SetDataSource(DS)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_HeaderGroupBy.ParameterFieldName, ParamHeaderGroupBy)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_Judul.ParameterFieldName, ParamJudulGroup)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_StartDate.ParameterFieldName, paramStartDate)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_EndDate.ParameterFieldName, paramEndDate)
            Me.crv.ReportSource = ReportDetail
            Me.crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            'DTACetak_Header.Dispose()
            DTACetak_Item.Dispose()
            'TBLCetak_Header.Dispose()
            TBLCetak_Item.Dispose()
            'TBLHeader.Dispose()
            TBLItem.Dispose()
        End Try
    End Sub

    Private Sub RboProduct_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RboProduct.CheckedChanged
        'Try
        '    If Me.RboProduct.Checked = True Then
        '        Me.cboCodeNo.Enabled = True
        '    Else
        '        Me.cboCodeNo.Enabled = False
        '        Me.cboCodeNo.SelectedIndex = 0
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub
End Class