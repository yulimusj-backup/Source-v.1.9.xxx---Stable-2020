Option Explicit On
Option Strict Off

Public Class FrmCetakQuotation
    Public Primary As Double = -1
    Public FrmParent As FrmTransaksiQuotation
    Public ListPrimary As List(Of Double)
    Public Batch As Boolean = False

    Dim ObjCompany As New Core_Module.ClsMasterCompanyData

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Me.Close()
            Me.cboLanguage.SelectedIndex = 0
            Me.cboVersion.SelectedIndex = 0
            Me.NonAktifCRV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmCetakQuotation_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.cbocontrol()
            Me.NonAktifCRV()
            Me.BtnPreview.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cbocontrol()
        Try
            Me.cboLanguage.Items.Clear()
            Me.cboVersion.Items.Clear()

            Me.cboLanguage.Items.Add("English")
            Me.cboLanguage.Items.Add("Indonesian")
            Me.cboLanguage.SelectedIndex = 1

            Me.cboVersion.Items.Add("Short")
            Me.cboVersion.Items.Add("Long")
            Me.cboVersion.SelectedIndex = 0

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cetak()
        Dim obj As New ClsTransaksiQuotation
        Dim CompanyCity As String = ""

        Try
            CompanyCity = Me.ObjCompany.GetCompany_City
            If CompanyCity <> "" Then
                CompanyCity = CompanyCity + ", "
            End If
            If Me.Batch = False Then
                If Me.cboLanguage.SelectedIndex = 0 And Me.cboVersion.SelectedIndex = 0 Then
                    obj.CetakQuotationVersiEnglishPendek(crv1, Me.Primary, CompanyCity)
                    Me.crv1.Visible = True
                    Me.TabControl1.Visible = False
                ElseIf Me.cboLanguage.SelectedIndex = 0 And Me.cboVersion.SelectedIndex = 1 Then
                    obj.CetakQuotationVersiEnglishPanjang(crv2, Me.Primary, CompanyCity)
                    obj.CetakQuotationVersiDetail(CrvDetail, Me.Primary)
                    Me.TabControl1.Visible = True
                    Me.crv1.Visible = False
                ElseIf Me.cboLanguage.SelectedIndex = 1 And Me.cboVersion.SelectedIndex = 0 Then
                    obj.CetakQuotationVersiIndoPendek(crv1, Me.Primary, CompanyCity)
                    Me.crv1.Visible = True
                    Me.TabControl1.Visible = False
                ElseIf Me.cboLanguage.SelectedIndex = 1 And Me.cboVersion.SelectedIndex = 1 Then
                    obj.CetakQuotationVersiIndoPanjang(crv2, Me.Primary, CompanyCity)
                    obj.CetakQuotationVersiDetail(CrvDetail, Me.Primary)
                    Me.TabControl1.Visible = True
                    Me.crv1.Visible = False
                Else
                    Me.NonAktifCRV()
                End If
            Else
                If Me.cboLanguage.SelectedIndex = 0 And Me.cboVersion.SelectedIndex = 0 Then
                    obj.CetakQuotation_Batch(Me.ListPrimary, Me.cboLanguage.SelectedIndex, Me.cboVersion.SelectedIndex, _
                    False, crv1, CrvDetail)
                    Me.crv1.Visible = True
                    Me.TabControl1.Visible = False
                ElseIf Me.cboLanguage.SelectedIndex = 0 And Me.cboVersion.SelectedIndex = 1 Then
                    obj.CetakQuotation_Batch(Me.ListPrimary, Me.cboLanguage.SelectedIndex, Me.cboVersion.SelectedIndex, _
                    False, crv2, CrvDetail)
                    Me.TabControl1.Visible = True
                    Me.crv1.Visible = False
                ElseIf Me.cboLanguage.SelectedIndex = 1 And Me.cboVersion.SelectedIndex = 0 Then
                    obj.CetakQuotation_Batch(Me.ListPrimary, Me.cboLanguage.SelectedIndex, Me.cboVersion.SelectedIndex, _
                    False, crv1, CrvDetail)
                    Me.crv1.Visible = True
                    Me.TabControl1.Visible = False
                ElseIf Me.cboLanguage.SelectedIndex = 1 And Me.cboVersion.SelectedIndex = 1 Then
                    obj.CetakQuotation_Batch(Me.ListPrimary, Me.cboLanguage.SelectedIndex, Me.cboVersion.SelectedIndex, _
                    False, crv2, CrvDetail)
                    Me.TabControl1.Visible = True
                    Me.crv1.Visible = False
                Else
                    Me.NonAktifCRV()
                End If
            End If
            
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Sub

    Private Sub BtnPreview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnPreview.Click
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.cetak()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifCRV()
        Try
            Me.crv1.Visible = False
            Me.TabControl1.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TabControl1.SelectedIndex = 0 Then
                Me.crv2.DisplayStatusBar = True
                Me.crv2.DisplayToolbar = True
            Else
                Me.CrvDetail.DisplayToolbar = True
                Me.CrvDetail.DisplayStatusBar = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class