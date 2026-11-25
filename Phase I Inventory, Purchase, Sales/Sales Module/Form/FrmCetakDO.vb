Option Explicit On
Option Strict Off

Public Class FrmCetakDO
    Public Primary As Double = -1
    Public UserCode As String = ""
    Public SudahAdaStockOut As Boolean = False
    Dim Print As Boolean = False
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting

    Dim DTADO As New DseTransaksiDOTableAdapters.tbltrans_do5TableAdapter
    Dim TBLDO As New DseTransaksiDO.tbltrans_do5DataTable
    Dim rowDO As DseTransaksiDO.tbltrans_do5Row

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmCetakDO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.SudahAdaStockOut = False And Me.Print = True Then
                If MsgBox("Is DO printed correctly?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DTADO.Update(1, Me.UserCode, Now, Me.Primary)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmCetakDO_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Obj As New CustomSMESS.ClsTransaksiDO
        Dim PrintedBy As String = ""
        Dim PrintTime As String = ""
        Dim ObjDO As New ClsTransaksiDO
        Dim CetakCustom As String = ""

        Try
            Me.SudahAdaStockOut = ObjDO.ApakahSudahAdaStockOutSales(Me.Primary)
            CetakCustom = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomDO_Bool)
            If CetakCustom = "" Or CetakCustom = "0" Then
                If Me.SudahAdaStockOut = True Then
                    ObjDO.CetakDOCMSI_Panjanga(Me.crv1, Me.Primary, Me.UserCode)
                Else
                    If Me.DTADO.FillBy_PrimaryDO(TBLDO, Me.Primary) > 0 Then
                        rowDO = TBLDO.Rows(0)
                        If rowDO.SudahCetakDOPendek = 0 Then
                            ObjDO.CetakDOCMSI_Pendeka(Me.crv1, Me.Primary)
                            Me.Print = True
                        Else
                            PrintedBy = rowDO.User_Cetak_DO_Pendek
                            PrintTime = Format(rowDO.Time_Cetak_DO_Pendek, "dd MMM yyyy HH:mm:ss")

                            If MsgBox("This DO has been printed by: " + PrintedBy + " on " + PrintTime + ". Continue Printing?", _
                            MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                ObjDO.CetakDOCMSI_Pendeka(Me.crv1, Me.Primary)
                                Me.Print = True
                            Else
                                Me.Close()
                            End If
                        End If
                    Else
                        ObjDO.CetakDOCMSI_Pendeka(Me.crv1, Me.Primary)
                        Me.Print = True
                    End If
                End If
            Else
                If Obj.CetakDO(Me.crv1, Me.Primary, Me.UserCode) = False Then
                    Me.Close()
                Else
                    Me.Print = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
        End Try
    End Sub
End Class