Option Explicit On
Option Strict Off

Public Class frmLaporanStockByBatchNumberNExpiryDate
    Public BackColorAll As Color = BackColorAll
    Dim JenisTgl As String
    Dim BrandKah As Boolean = False
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Sub cboPeriod_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCode.SelectedIndexChanged
        Me.KosongkanDetail()
        Me.SesuaikanBrandOrProduct()
    End Sub

    Private Sub KosongkanDetail()
        Try
            Me.txtPrimary.Text = "-1"
            Me.txtBrandCodeOrCodeNumber.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SesuaikanBrandOrProduct()
        Try
            If Me.cboCode.SelectedIndex = 0 Then
                'Product Code
                Me.lbBrandCodeOrCodeNumber.Text = "Product Code"
                Me.BrandKah = False
            Else
                'Brand Code
                Me.lbBrandCodeOrCodeNumber.Text = "Brand Code"
                Me.BrandKah = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SesuaikanTgl()
        Try
            If Me.cboTgl.SelectedIndex = 0 Then
                'After
                Me.lbTgl1.Visible = True
                Me.dtpTgl1.Visible = True
                Me.lbTgl1.Text = "After"

                Me.lbTgl2.Visible = False
                Me.dtpTgl2.Visible = False
                Me.JenisTgl = "after"
            ElseIf Me.cboTgl.SelectedIndex = 1 Then
                'Before
                Me.lbTgl1.Visible = True
                Me.dtpTgl1.Visible = True
                Me.lbTgl1.Text = "Before"

                Me.lbTgl2.Visible = False
                Me.dtpTgl2.Visible = False
                Me.JenisTgl = "before"
            ElseIf Me.cboTgl.SelectedIndex = 2 Then
                'Between
                Me.lbTgl1.Visible = True
                Me.dtpTgl1.Visible = True
                Me.lbTgl1.Text = "From"

                Me.lbTgl2.Visible = True
                Me.lbTgl2.Text = "thru"
                Me.dtpTgl2.Visible = True
                Me.JenisTgl = "between"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmLaporanStockByBatchNumberNExpiryDate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.cboCode.SelectedIndex = 0
            Me.cboTgl.SelectedIndex = 0
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cetak()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim report As New crpLaporanStockByBatchNumberNExpiryDate
        Dim DS As New DseLaporanStockByBatchNExpiryDate
        Dim DTA As New DseLaporanStockByBatchNExpiryDateTableAdapters.laporan_stockbybatchnexpirydate_view0TableAdapter
        Dim BrandCode As String = ""
        Dim CodeNumber As String = ""
        Try
            If Me.cboCode.SelectedIndex = 0 Then
                'Product Code
                BrandCode = DTA.ScalarQuery_getBrandCodeByPrimaryProduct(Me.txtPrimary.Text)
                If String.IsNullOrEmpty(BrandCode) = True Then
                    BrandCode = "Unidentified Brand Code"
                End If
                CodeNumber = Me.txtBrandCodeOrCodeNumber.Text
                If Me.cboTgl.SelectedIndex = 0 Then
                    'after
                    DTA.FillBy_PrimaryProductNAfterDate(DS.laporan_stockbybatchnexpirydate_view0, Me.txtPrimary.Text, Me.dtpTgl1.Value)
                ElseIf Me.cboTgl.SelectedIndex = 1 Then
                    'before
                    DTA.FillBy_PrimaryProductNBeforeDate(DS.laporan_stockbybatchnexpirydate_view0, Me.txtPrimary.Text, Me.dtpTgl1.Value)
                ElseIf Me.cboTgl.SelectedIndex = 2 Then
                    'between
                    DTA.FillBy_PrimaryProductNTglDariSampai(DS.laporan_stockbybatchnexpirydate_view0, Me.txtPrimary.Text, Me.dtpTgl1.Value, Me.dtpTgl2.Value)
                End If
            ElseIf Me.cboCode.SelectedIndex = 1 Then
                'Brand Code
                If Me.cboTgl.SelectedIndex = 0 Then
                    'after
                    DTA.FillBy_PrimaryBrandNAfterDate(DS.laporan_stockbybatchnexpirydate_view0, Me.txtPrimary.Text, Me.dtpTgl1.Value)
                ElseIf Me.cboTgl.SelectedIndex = 1 Then
                    'before
                    DTA.FillBy_PrimaryBrandNBeforeDate(DS.laporan_stockbybatchnexpirydate_view0, Me.txtPrimary.Text, Me.dtpTgl1.Value)
                ElseIf Me.cboTgl.SelectedIndex = 2 Then
                    'between
                    DTA.FillBy_PrimaryBrandNTglDariSampai(DS.laporan_stockbybatchnexpirydate_view0, Me.txtPrimary.Text, Me.dtpTgl1.Value, Me.dtpTgl2.Value)
                End If
                BrandCode = Me.txtBrandCodeOrCodeNumber.Text
                CodeNumber = ""
            End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_BrandCode.ParameterFieldName, BrandCode)
            report.SetParameterValue(report.Parameter_CodeNumber.ParameterFieldName, CodeNumber)
            report.SetParameterValue(report.Parameter_BrandKah.ParameterFieldName, Me.BrandKah)
            report.SetParameterValue(report.Parameter_JenisTgl.ParameterFieldName, Me.JenisTgl)
            report.SetParameterValue(report.Parameter_TglPertama.ParameterFieldName, Me.dtpTgl1.Value)
            report.SetParameterValue(report.Parameter_TglKedua.ParameterFieldName, Me.dtpTgl2.Value)

            Me.crv.ReportSource = report
            Me.crv.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Try
            If Me.cboCode.SelectedIndex = 0 Then
                'Product Code
                Dim i As New Core_Module.frmCariMasterProduct
                If Me.txtBrandCodeOrCodeNumber.Text <> "" Then
                    i.txtCari.Text = Me.txtBrandCodeOrCodeNumber.Text
                End If

                i.ShowDialog()
                If i.PrimaryProduct > 0 Then
                    Me.txtPrimary.Text = i.PrimaryProduct
                    Dim objproduct As New Core_Module.clsMasterProduct
                    Me.txtBrandCodeOrCodeNumber.Text = objproduct.GetProductCode(Me.txtPrimary.Text)
                    objproduct = Nothing
                    Me.dtpTgl1.Focus()
                Else
                    Me.txtPrimary.Text = "-1"
                    Me.txtBrandCodeOrCodeNumber.Text = ""
                    Me.txtBrandCodeOrCodeNumber.Focus()
                End If
            ElseIf Me.cboCode.SelectedIndex = 1 Then
                'Brand Code
                Dim i As New Core_Module.FrmCariMasterProductGroup
                i.ProductGroupLevel = 2
                If Me.txtBrandCodeOrCodeNumber.Text <> "" Then
                    i.txtCari.Text = Me.txtBrandCodeOrCodeNumber.Text
                End If

                i.ShowDialog()
                If i.PrimaryProductGroup > 0 Then
                    Me.txtPrimary.Text = i.PrimaryProductGroup
                    Dim objproductGroup As New Core_Module.ClsMasterProductGroup
                    Me.txtBrandCodeOrCodeNumber.Text = objproductGroup.GetGroupCode(Me.txtPrimary.Text)
                    objproductGroup = Nothing
                    Me.dtpTgl1.Focus()
                Else
                    Me.txtPrimary.Text = "-1"
                    Me.txtBrandCodeOrCodeNumber.Text = ""
                    Me.txtBrandCodeOrCodeNumber.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBrandCodeOrCodeNumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBrandCodeOrCodeNumber.GotFocus
        Try
            Me.txtBrandCodeOrCodeNumber.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBrandCodeOrCodeNumber_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBrandCodeOrCodeNumber.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                If Me.cboCode.SelectedIndex = 0 Then
                    Dim obj As New Core_Module.clsMasterProduct
                    Dim primary As Int16 = obj.GetPrimary(Me.txtBrandCodeOrCodeNumber.Text)
                    obj = Nothing
                    If primary = -1 Then
                        Me.btnSearch.PerformClick()
                    Else
                        Me.txtPrimary.Text = CStr(primary)
                    End If
                ElseIf Me.cboCode.SelectedIndex = 1 Then
                    Dim obj As New Core_Module.ClsMasterProductGroup
                    Dim primary As Int16 = obj.GetPrimary(Me.txtBrandCodeOrCodeNumber.Text, 2)
                    obj = Nothing
                    If primary = -1 Then
                        Me.btnSearch.PerformClick()
                    Else
                        Me.txtPrimary.Text = CStr(primary)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBrandCodeOrCodeNumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBrandCodeOrCodeNumber.Leave
        Try
            If String.IsNullOrEmpty(Me.txtBrandCodeOrCodeNumber.Text) = False Then
                If Me.cboCode.SelectedIndex = 0 Then
                    Dim obj As New Core_Module.clsMasterProduct
                    Dim primary As Int16 = obj.GetPrimary(Me.txtBrandCodeOrCodeNumber.Text)
                    obj = Nothing
                    If primary = -1 Then
                        Me.btnSearch.PerformClick()
                    Else
                        Me.txtPrimary.Text = CStr(primary)
                    End If
                ElseIf Me.cboCode.SelectedIndex = 1 Then
                    Dim obj As New Core_Module.ClsMasterProductGroup
                    Dim primary As Int16 = obj.GetPrimary(Me.txtBrandCodeOrCodeNumber.Text, 2)
                    obj = Nothing
                    If primary = -1 Then
                        Me.btnSearch.PerformClick()
                    Else
                        Me.txtPrimary.Text = CStr(primary)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim EndLicenseDate As Date = #1/1/1950#

        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpTgl1.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTgl1.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTgl2.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTgl2.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If String.IsNullOrEmpty(Me.txtBrandCodeOrCodeNumber.Text) = False Then
                Me.Cetak()
            Else
                If Me.cboCode.SelectedIndex = 0 Then
                    MsgBox("Please insert product code.", MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Please insert brand code.", MsgBoxStyle.Information, "Information")
                    Me.txtBrandCodeOrCodeNumber.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboTgl_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTgl.SelectedIndexChanged
        Me.SesuaikanTgl()
    End Sub

    Private Sub txtBrandCodeOrCodeNumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBrandCodeOrCodeNumber.LostFocus
        Try
            Me.txtBrandCodeOrCodeNumber.BackColor = Me.txtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class