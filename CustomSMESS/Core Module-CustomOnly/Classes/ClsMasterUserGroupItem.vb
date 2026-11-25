Option Explicit On
Option Strict Off

Public Class ClsMasterUserGroupItem
    Public Function IsiDGV(ByVal vTrv As Windows.Forms.TreeView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseMasterUserGroupTableAdapters.tblmaster_usergroup_itemTableAdapter
        Dim TBL As New DseMasterUserGroup.tblmaster_usergroup_itemDataTable
        Dim Nd, NdItem As Windows.Forms.TreeNode
        Dim Filter As String = ""
        Dim rowTBL As DseMasterUserGroup.tblmaster_usergroup_itemRow

        Try
            'isikan tabel
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                For Each Nd In vTrv.Nodes
                    Filter = "[Kode Hak Akses] = " + Nd.Name + ""

                    TBL.DefaultView.RowFilter = Filter
                    If TBL.DefaultView.Count = 0 Then
                        Nd.Checked = False
                    Else
                        rowTBL = TBL.DefaultView(0).Row
                        If rowTBL.Checked = 0 Then
                            Nd.Checked = False
                        Else
                            Nd.Checked = True
                        End If
                    End If

                    For Each NdItem In Nd.Nodes
                        Me.SetTreeView(NdItem, TBL)
                    Next
                Next
            Else
                For Each Nd In vTrv.Nodes
                    Me.KosongkanTreeView(Nd)
                Next
            End If

            'tandai status pada check box pada node yang tidak secara 
            'langsung terhubung dengan basis data
            Me.Set_Status_Node(vTrv)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
    End Function

    Public Sub Set_Status_Node(ByVal vTrv As TreeView)
        Dim ND As TreeNode
        Try
            For Each ND In vTrv.Nodes
                Me.Set_Status_Node(ND)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function Set_Status_Node(ByVal ND_Parent As TreeNode) As Double
        Dim ND_child As TreeNode
        Dim Jml_True, Jml_False, Dummy As Double

        Try
            If ND_Parent.Nodes.Count > 0 Then
                For Each ND_child In ND_Parent.Nodes
                    Dummy = Me.Set_Status_Node(ND_child)

                    If Dummy = -1 Then
                        Jml_False += 1
                    Else
                        Jml_True += 1
                    End If
                Next

                If Jml_True > 0 And Jml_False = 0 Then
                    ND_Parent.Checked = True
                    Return 1
                Else
                    ND_Parent.Checked = False
                    Return -1
                End If
            Else
                If ND_Parent.Checked = True Then
                    Return 1
                Else
                    Return -1
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub KosongkanTreeView(ByVal Nd As Windows.Forms.TreeNode)
        Dim NdItem As Windows.Forms.TreeNode

        Try
            Nd.Checked = False

            If Nd.GetNodeCount(True) > 0 Then
                For Each NdItem In Nd.Nodes
                    Me.KosongkanTreeView(NdItem)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetTreeView(ByVal Nd As Windows.Forms.TreeNode, _
    ByVal TBL As DseMasterUserGroup.tblmaster_usergroup_itemDataTable)
        Dim Filter As String = ""
        Dim rowTBL As DseMasterUserGroup.tblmaster_usergroup_itemRow
        Dim NdItem As TreeNode

        Try
            Filter = "[Kode Hak Akses] = " + Nd.Name + ""

            TBL.DefaultView.RowFilter = Filter
            If TBL.DefaultView.Count = 0 Then
                Nd.Checked = False
            Else
                rowTBL = TBL.DefaultView(0).Row
                If rowTBL.Checked = 0 Then
                    Nd.Checked = False
                Else
                    Nd.Checked = True
                End If
            End If

            For Each NdItem In Nd.Nodes
                Me.SetTreeView(NdItem, TBL)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function Simpan(ByVal trv As TreeView, _
    ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseMasterUserGroupTableAdapters.tblmaster_usergroup_itemTableAdapter
        Dim vNd, NDItem As Windows.Forms.TreeNode
        Dim CheckInt As Integer = 0

        Try
            For Each vNd In trv.Nodes
                If vNd.Name <> -1 Then
                    If vNd.Checked = True Then
                        CheckInt = 1
                    Else
                        CheckInt = 0
                    End If
                    DTA.Insert(vPrimaryTransaksi, vNd.Name, CheckInt)
                End If

                If vNd.GetNodeCount(True) > 0 Then
                    For Each NDItem In vNd.Nodes
                        Me.Simpan(DTA, NDItem, vPrimaryTransaksi)
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub Simpan(ByVal DTA As _
    DseMasterUserGroupTableAdapters.tblmaster_usergroup_itemTableAdapter, _
    ByVal ND As Windows.Forms.TreeNode, ByVal vPrimaryTransaksi As Double)
        Dim NDItem As Windows.Forms.TreeNode
        Dim CheckInt As Integer = 0

        Try
            If ND.Name <> -1 Then
                If ND.Checked = True Then
                    CheckInt = 1
                Else
                    CheckInt = 0
                End If
                DTA.Insert(vPrimaryTransaksi, ND.Name, CheckInt)
            End If

            If ND.GetNodeCount(True) > 0 Then
                For Each NDItem In ND.Nodes
                    Me.Simpan(DTA, NDItem, vPrimaryTransaksi)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Erro")
        End Try
    End Sub

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseMasterUserGroupTableAdapters.tblmaster_usergroup_itemTableAdapter
        Try
            'hapus data item
            DTA.Delete(vPrimaryTransaksi)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


End Class
