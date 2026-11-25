Option Explicit On
Option Strict Off

Public Class FrmSet_IO_Location
    Dim ObjWUGroup As New Core_Module.ClsMasterWorkUnitGroup
    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub btnSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSet.Click
        Dim DTA As New DseSet_IOTableAdapters.tblmaster_workunitTableAdapter
        Dim TBL As New DseSet_IO.tblmaster_workunitDataTable
        Dim row As DseSet_IO.tblmaster_workunitRow

        Try
            'kumpulkan wu dengan primary wu group ybs
            DTA.FillBy_PrimaryWUGroup(TBL, Me.TxtPrimaryWUGroup.Text)
            For Each row In TBL.Rows
                'untuk setiap wu, set primary input location, set primary output location
                DTA.UpdateQuery(Me.TxtPrimaryInputLocation.Text, Me.TxtPrimaryOutputLocation.Text, row.primary)
            Next
            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWUGroupCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtWUGroupCode.KeyDown
        Dim i As New Core_Module.FrmMasterWorkUnitGroup_Cari
        Dim Jml As Long = 0
        Dim PrimaryWUGroup As Double = -1

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjWUGroup.GetJml(Me.TxtWUGroupCode.Text)
                PrimaryWUGroup = Me.ObjWUGroup.GetPrimary(Me.TxtWUGroupCode.Text)

                If Jml = 1 Then
                    Me.TxtPrimaryWUGroup.Text = Me.ObjWUGroup.GetPrimary(Me.TxtWUGroupCode.Text)
                    Me.TxtWUGroupDesc.Text = Me.ObjWUGroup.GetDesc(Me.TxtPrimaryWUGroup.Text)
                    Me.TxtInputLocationCode.Focus()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtWUGroupCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtWUGroupCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryWorkUnitGroup > 0 Then
                        Me.TxtPrimaryWUGroup.Text = i.PrimaryWorkUnitGroup
                        Me.TxtWUGroupCode.Text = Me.ObjWUGroup.GetCode(Me.TxtPrimaryWUGroup.Text)
                        Me.TxtWUGroupDesc.Text = Me.ObjWUGroup.GetDesc(Me.TxtPrimaryWUGroup.Text)

                        Me.TxtInputLocationCode.Focus()
                    Else
                        Me.TxtWUGroupCode.Clear()
                        Me.TxtPrimaryWUGroup.Text = "-1"
                        Me.TxtWUGroupDesc.Text = ""
                        Me.TxtWUGroupCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtWUGroupCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWUGroupCode.Leave
        Dim i As New Core_Module.FrmMasterWorkUnitGroup_Cari
        Dim PrimaryWUGroup As Double = -1

        Try
            If Me.TxtWUGroupCode.Text <> "" Then
                PrimaryWUGroup = Me.ObjWUGroup.GetPrimary(Me.TxtWUGroupCode.Text)
                If Me.ObjWUGroup.GetJml(Me.TxtWUGroupCode.Text) = 1 Then
                    Me.TxtPrimaryWUGroup.Text = Me.ObjWUGroup.GetPrimary(Me.TxtWUGroupCode.Text)
                    Me.TxtWUGroupDesc.Text = Me.ObjWUGroup.GetDesc(Me.TxtPrimaryWUGroup.Text)
                    Me.TxtInputLocationCode.Focus()
                Else
                    If Me.TxtWUGroupCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtWUGroupCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryWorkUnitGroup > 0 Then
                        Me.TxtPrimaryWUGroup.Text = i.PrimaryWorkUnitGroup
                        Me.TxtWUGroupCode.Text = Me.ObjWUGroup.GetCode(Me.TxtPrimaryWUGroup.Text)
                        Me.TxtWUGroupDesc.Text = Me.ObjWUGroup.GetDesc(Me.TxtPrimaryWUGroup.Text)

                        Me.TxtInputLocationCode.Focus()
                    Else
                        Me.TxtWUGroupCode.Clear()
                        Me.TxtWUGroupDesc.Text = ""
                        Me.TxtPrimaryWUGroup.Text = "-1"
                        Me.TxtWUGroupCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryWUGroup.Text = "-1"
                Me.TxtWUGroupDesc.Clear()
                Me.TxtWUGroupDesc.Text = ""
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindWUGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindWUGroup.Click
        Dim i As New Core_Module.FrmMasterWorkUnitGroup_Cari

        Try
            If Me.TxtWUGroupCode.Text <> "" Then
                i.txtCari.Text = Me.TxtWUGroupCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryWorkUnitGroup > 0 Then
                Me.TxtPrimaryWUGroup.Text = i.PrimaryWorkUnitGroup
                Me.TxtWUGroupCode.Text = Me.ObjWUGroup.GetCode(Me.TxtPrimaryWUGroup.Text)
                Me.TxtWUGroupDesc.Text = Me.ObjWUGroup.GetDesc(Me.TxtPrimaryWUGroup.Text)

                Me.TxtInputLocationCode.Focus()
            Else
                Me.TxtWUGroupCode.Clear()
                Me.TxtPrimaryWUGroup.Text = "-1"
                Me.TxtWUGroupDesc.Text = ""
                Me.TxtWUGroupCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtInputLocationCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtInputLocationCode.KeyDown
        Dim i As New Core_Module.FrmCariMasterStorageArea
        Dim Jml As Long = 0
        Dim PrimaryWUGroup As Double = -1

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjLocation.GetJml_ByCode(Me.TxtInputLocationCode.Text)
                PrimaryWUGroup = Me.ObjLocation.GetPrimary(Me.TxtInputLocationCode.Text)

                If Jml = 1 Then
                    Me.TxtPrimaryInputLocation.Text = Me.ObjLocation.GetPrimary(Me.TxtInputLocationCode.Text)
                    Me.TxtInputLocationDesc.Text = Me.ObjLocation.GetDescription(Me.TxtPrimaryInputLocation.Text)
                    Me.TxtOutputLocationCode.Focus()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtInputLocationCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtInputLocationCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.TxtPrimaryInputLocation.Text = i.PrimaryStorageArea
                        Me.TxtInputLocationCode.Text = Me.ObjLocation.GetStorageAreaCode(Me.TxtPrimaryInputLocation.Text)
                        Me.TxtInputLocationDesc.Text = Me.ObjLocation.GetDescription(Me.TxtPrimaryInputLocation.Text)

                        Me.TxtOutputLocationCode.Focus()
                    Else
                        Me.TxtInputLocationCode.Clear()
                        Me.TxtPrimaryInputLocation.Text = "-1"
                        Me.TxtInputLocationDesc.Text = ""
                        Me.TxtOutputLocationCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtInputLocationCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInputLocationCode.Leave
        Dim i As New Core_Module.FrmCariMasterStorageArea
        Dim PrimaryWUGroup As Double = -1

        Try
            If Me.TxtInputLocationCode.Text <> "" Then
                PrimaryWUGroup = Me.ObjLocation.GetPrimary(Me.TxtInputLocationCode.Text)
                If Me.ObjLocation.GetJml_ByCode(Me.TxtInputLocationCode.Text) = 1 Then
                    Me.TxtPrimaryInputLocation.Text = Me.ObjLocation.GetPrimary(Me.TxtInputLocationCode.Text)
                    Me.TxtInputLocationDesc.Text = Me.ObjLocation.GetDescription(Me.TxtPrimaryInputLocation.Text)

                    Me.TxtOutputLocationCode.Focus()
                Else
                    If Me.TxtInputLocationCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtInputLocationCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.TxtPrimaryInputLocation.Text = i.PrimaryStorageArea
                        Me.TxtInputLocationCode.Text = Me.ObjLocation.GetStorageAreaCode(Me.TxtPrimaryInputLocation.Text)
                        Me.TxtInputLocationDesc.Text = Me.ObjLocation.GetDescription(Me.TxtPrimaryInputLocation.Text)

                        Me.TxtOutputLocationCode.Focus()
                    Else
                        Me.TxtInputLocationCode.Clear()
                        Me.TxtInputLocationDesc.Text = ""
                        Me.TxtPrimaryInputLocation.Text = "-1"
                        Me.TxtOutputLocationCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryInputLocation.Text = "-1"
                Me.TxtInputLocationDesc.Text = ""
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindInputLocation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindInputLocation.Click
        Dim i As New Core_Module.FrmCariMasterStorageArea

        Try
            If Me.TxtInputLocationCode.Text <> "" Then
                i.txtCari.Text = Me.TxtInputLocationCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryStorageArea > 0 Then
                Me.TxtPrimaryInputLocation.Text = i.PrimaryStorageArea
                Me.TxtInputLocationCode.Text = Me.ObjLocation.GetStorageAreaCode(Me.TxtPrimaryInputLocation.Text)
                Me.TxtInputLocationDesc.Text = Me.ObjLocation.GetDescription(Me.TxtPrimaryInputLocation.Text)

                Me.TxtOutputLocationCode.Focus()
            Else
                Me.TxtInputLocationCode.Clear()
                Me.TxtPrimaryInputLocation.Text = "-1"
                Me.TxtInputLocationDesc.Text = ""
                Me.TxtOutputLocationCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtOutputLocationCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtOutputLocationCode.KeyDown
        Dim i As New Core_Module.FrmCariMasterStorageArea
        Dim Jml As Long = 0
        Dim PrimaryWUGroup As Double = -1

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjLocation.GetJml_ByCode(Me.TxtOutputLocationCode.Text)
                PrimaryWUGroup = Me.ObjLocation.GetPrimary(Me.TxtOutputLocationCode.Text)

                If Jml = 1 Then
                    Me.TxtPrimaryOutputLocation.Text = Me.ObjLocation.GetPrimary(Me.TxtOutputLocationCode.Text)
                    Me.TxtOutputLocationDesc.Text = Me.ObjLocation.GetDescription(Me.TxtPrimaryOutputLocation.Text)
                    Me.btnSet.Focus()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtOutputLocationCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtOutputLocationCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.TxtPrimaryOutputLocation.Text = i.PrimaryStorageArea
                        Me.TxtOutputLocationCode.Text = Me.ObjLocation.GetStorageAreaCode(Me.TxtPrimaryOutputLocation.Text)
                        Me.TxtOutputLocationDesc.Text = Me.ObjLocation.GetDescription(Me.TxtPrimaryOutputLocation.Text)

                        Me.btnSet.Focus()
                    Else
                        Me.TxtOutputLocationCode.Clear()
                        Me.TxtPrimaryOutputLocation.Text = "-1"
                        Me.TxtOutputLocationDesc.Text = ""
                        Me.btnSet.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtOutputLocationCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtOutputLocationCode.Leave
        Dim i As New Core_Module.FrmCariMasterStorageArea
        Dim PrimaryWUGroup As Double = -1

        Try
            If Me.TxtOutputLocationCode.Text <> "" Then
                PrimaryWUGroup = Me.ObjLocation.GetPrimary(Me.TxtOutputLocationCode.Text)
                If Me.ObjLocation.GetJml_ByCode(Me.TxtOutputLocationCode.Text) = 1 Then
                    Me.TxtPrimaryOutputLocation.Text = Me.ObjLocation.GetPrimary(Me.TxtOutputLocationCode.Text)
                    Me.TxtOutputLocationDesc.Text = Me.ObjLocation.GetDescription(Me.TxtPrimaryOutputLocation.Text)

                    Me.btnSet.Focus()
                Else
                    If Me.TxtOutputLocationCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtOutputLocationCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.TxtPrimaryOutputLocation.Text = i.PrimaryStorageArea
                        Me.TxtOutputLocationCode.Text = Me.ObjLocation.GetStorageAreaCode(Me.TxtPrimaryOutputLocation.Text)
                        Me.TxtOutputLocationDesc.Text = Me.ObjLocation.GetDescription(Me.TxtPrimaryOutputLocation.Text)

                        Me.btnSet.Focus()
                    Else
                        Me.TxtOutputLocationCode.Clear()
                        Me.TxtOutputLocationDesc.Text = ""
                        Me.TxtPrimaryOutputLocation.Text = "-1"
                        Me.btnSet.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryOutputLocation.Text = "-1"
                Me.TxtOutputLocationDesc.Text = ""
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindOutputLocation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindOutputLocation.Click
        Dim i As New Core_Module.FrmCariMasterStorageArea

        Try
            If Me.TxtOutputLocationCode.Text <> "" Then
                i.txtCari.Text = Me.TxtOutputLocationCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryStorageArea > 0 Then
                Me.TxtPrimaryOutputLocation.Text = i.PrimaryStorageArea
                Me.TxtOutputLocationCode.Text = Me.ObjLocation.GetStorageAreaCode(Me.TxtPrimaryOutputLocation.Text)
                Me.TxtOutputLocationDesc.Text = Me.ObjLocation.GetDescription(Me.TxtPrimaryOutputLocation.Text)

                Me.btnSet.Focus()
            Else
                Me.TxtOutputLocationCode.Clear()
                Me.TxtPrimaryOutputLocation.Text = "-1"
                Me.TxtOutputLocationDesc.Text = ""
                Me.btnSet.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub
End Class