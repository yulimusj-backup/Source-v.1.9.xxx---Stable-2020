Option Explicit On
Option Strict Off
Public Class FrmManajemenHakAkses2
    Public MainMenuStrip1 As MenuStrip
    ''Public TreeNode As TreeNode
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Dim ObjHak As New ClsMasterUserGroup
    Dim ObjHak_Item As New ClsMasterUserGroupItem
    Dim baru As Boolean
    Dim TBL As New DSUser.tblmasterhakaksesDataTable
    Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter
    Dim Nd As Windows.Forms.TreeNode
    Public ObjCD As New ClsCekData
    Dim Aktif As Boolean = False
    Dim Berubah As Boolean = False

    Private Sub FrmManajemenHakAkses2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.trvAccess.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjHak.SimpanBaru(Me.trvAccess, Me.UserCode) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                ''Me.NonAktifkanDetail()
                                ''Me.btnFind.Focus()
                                Me.btnOK.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjHak.SimpanPerubahan(Me.trvAccess, Me.UserCode, Me.TxtPrimaryGroup.Text) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                ''Me.NonAktifkanDetail()
                                ''Me.btnEdit.Focus()
                                Me.btnOK.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmManajemenHakAkses2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'isikan tree view sesuai dengan menu strip yang diberikan
            DTA.Fill(TBL)
            Me.FillTree(MainMenuStrip1)
            Me.NonAktifkandetail()

            'kosongkan TextDescription
            Me.TxtAccesDesc.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FillDescription(ByVal TN As TreeNode, ByVal tulis As TextBox)
        Dim Obj As Object
        ''Dim tx As Windows.Forms.TreeNode
        ''Dim Tulis As Windows.Forms.TextBox
        Try
            ''Buat
            If Me.trvAccess.Nodes.Count <> -1 Then
                For Each Obj In TN.Nodes
                    TN = TryCast(Obj, Windows.Forms.TreeNode)
                    If IsNothing(TN) = False Then
                        If TN.Checked = True Then
                            tulis.Text = Me.GetDescription(TN)
                        End If
                    End If
                Next
            End If

            'buat node
            ''If Me.trvAccess.Nodes.Count <> -1 Then
            ''For Each TN In Me.trvAccess.Nodes
            ''If TN.Checked = True Then
            ''Me.GetDescription(TN)
            ''End If
            ''Next
            ''End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetDescription(ByVal i As TreeNode) As String
        '' Dim ds As New DseMasterUserGroup
        ''Dim row As DseMasterUserGroup.tblmasterhakaksesRow
        Dim DTA As New DseMasterUserGroupTableAdapters.tblmasterhakaksesTableAdapter
        Try
            '' For Each row In ds.tblmasterhakakses.Rows
            ''Me.TxtAccesDesc.Text = row.Description
            ''Next
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDescription(i.Name)) = False Then
                Return DTA.ScalarQuery_GetDescription(i.Name)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub FillTree(ByVal MS As MenuStrip)
        Dim Obj As Object
        Dim Mn As Windows.Forms.ToolStripMenuItem

        Try
            'kosongkan nodes
            Me.trvAccess.Nodes.Clear()

            'buat node
            For Each Obj In MS.Items
                Mn = TryCast(Obj, Windows.Forms.ToolStripMenuItem)
                If IsNothing(Mn) = False Then
                    Nd = New Windows.Forms.TreeNode
                    Me.trvAccess.Nodes.Add(Nd)
                    Nd.Text = Mn.Text
                    Nd.Name = Me.GetKodeHakAkses(Mn.Name.Substring(0, Mn.Name.Length - 1)) '--> kode disimpan disini

                    'kalo menu item ada anaknya, maka tambahkan juga anaknya ke dalam tree view
                    If Mn.DropDownItems.Count > 0 Then
                        Me.FillItem(Nd, Mn)
                    End If
                End If
            Next

            'Bersihkan tanda dan 
            For Each Nd In Me.trvAccess.Nodes
                Me.Bersihkan_Tanda_Dan(Nd)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FillItem(ByVal Nd As Windows.Forms.TreeNode, ByVal Mn As Windows.Forms.ToolStripMenuItem)
        Dim Obj As Object
        Dim MnItem As Windows.Forms.ToolStripMenuItem
        Dim NdItem As Windows.Forms.TreeNode
        Try
            For Each Obj In Mn.DropDownItems
                MnItem = TryCast(Obj, Windows.Forms.ToolStripMenuItem)
                If IsNothing(MnItem) = False Then
                    If MnItem.Name.Length <> 0 Then
                        NdItem = New Windows.Forms.TreeNode
                        NdItem.Text = MnItem.Text
                        NdItem.Name = Me.GetKodeHakAkses(MnItem.Name.Substring(0, MnItem.Name.Length - 1)) '--> kode disimpan disini
                        Nd.Nodes.Add(NdItem)
                        'kalo menu ada anaknya maka tambahkan juga anaknya ke dalam tree view
                        If MnItem.DropDownItems.Count > 0 Then
                            Me.FillItem(NdItem, MnItem)
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Bersihkan_Tanda_Dan(ByVal ND As TreeNode)
        Dim ND1 As Windows.Forms.TreeNode
        Dim StartIndex, Index_Dan, Index_Dan2 As Integer
        Try
            StartIndex = 0
            Index_Dan = ND.Text.IndexOf("&", StartIndex)
            While Index_Dan <> -1 And StartIndex <= ND.Text.Length - 1
                Index_Dan2 = ND.Text.IndexOf("&&", StartIndex)
                If Index_Dan2 <> Index_Dan Then
                    'bukan tanda dan sesungguhnya
                    StartIndex = Index_Dan
                Else
                    StartIndex = Index_Dan + 1
                End If
                If Index_Dan = 0 Then
                    ND.Text = ND.Text.Substring(1)
                Else
                    ND.Text = ND.Text.Substring(0, Index_Dan) + ND.Text.Substring(Index_Dan + 1)
                End If
                Index_Dan = ND.Text.IndexOf("&", StartIndex)
            End While

            For Each ND1 In ND.Nodes
                Me.Bersihkan_Tanda_Dan(ND1)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Tambahan untuk simpan node
    Private Function GetKodeHakAkses(ByVal vNamaHakAkses As String) As Double
        Dim Filter As String = ""
        Dim row As DSUser.tblmasterhakaksesRow

        Try
            Filter = "Nama ='" + vNamaHakAkses + "'"
            Me.TBL.DefaultView.RowFilter = Filter
            If Me.TBL.DefaultView.Count > 0 Then
                row = Me.TBL.DefaultView(0).Row
                Return row.Kode
            End If

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Aktifkandetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryGroup.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.trvAccess.Enabled = True
            Me.trvAccess.BackColor = Color.White
            Me.TxtAccesDesc.BackColor = Me.TxtReadonly.BackColor
            Me.TxtAccesDesc.Enabled = True
            Me.Aktif = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkandetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.trvAccess.Enabled = False
            Me.TxtAccesDesc.Enabled = False
            Me.trvAccess.BackColor = Me.TxtReadonly.BackColor
            Me.TxtAccesDesc.BackColor = Me.TxtReadonly.BackColor
            Me.Aktif = False
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.TxtPrimaryGroup.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.TxtAccesDesc.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariMasterHakAkses2

        Try
            ij.ShowDialog()
            If Not ij.PrimaryLPB = -1 Then
                If Me.ObjHak.CariDetail(ij.PrimaryLPB, Me.trvAccess) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.TxtGroupCode.Focus()
            Me.TxtAccesDesc.Clear()
            Me.ObjHak.ResetObject(Me.trvAccess)
            Me.TxtGroupCode.Text = ""
            Me.Aktifkandetail()

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KetentuanNode()
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtGroupCode.Text = "" Then
                Me.NonAktifkandetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjHak.CariDetail(Me.TxtPrimaryGroup.Text, Me.trvAccess)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkandetail()
                Me.btnNew.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtGroupCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtGroupCode.Text <> "" Then
                If Me.ObjHak.ApakahBolehHapus(Me.TxtPrimaryGroup.Text, True, UserCode, True) = False Then
                    If Me.ObjHak.GetCreatedBy(Me.TxtPrimaryGroup.Text) = Me.UserCode Then
                        MsgBox("You cannot delete this data, there are user using this data.", _
                          MsgBoxStyle.Information, "Information")
                    End If
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjHak.Hapus(Me.TxtPrimaryGroup.Text, UserCode, True) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtGroupCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtGroupCode.Text <> "" Then
                If Me.ObjHak.ApakahBolehEdit(Me.ObjHak.Primary, UserCode, True, True) = True Then
                    Me.baru = False
                    Me.ObjHak.CariDetail(Me.ObjHak.Primary, Me.trvAccess)
                    Me.SalinKeField()
                    Me.TxtGroupCode.Focus()
                    Me.Aktifkandetail()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjHak.SimpanBaru(Me.trvAccess, Me.UserCode) Then
                            MsgBox("Data saved. You have to restart the program so the changes can take effect.", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkandetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjHak.SimpanPerubahan(Me.trvAccess, Me.UserCode, CDbl(Me.TxtPrimaryGroup.Text)) Then
                            MsgBox("Data saved. You have to restart the program so the changes can take effect.", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkandetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Code Group")
            KoleksiTextBox.Add(Me.TxtGroupCode)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtDescription)

            If ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'kalo sampe cek berkaitan dengan timestamp apakah masih benar
            If Me.baru = False Then
                If Me.ObjHak.TimeStamp <> Me.ObjHak.GetTimestamp(Me.ObjHak.Primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjHak.CariDetail(Me.ObjHak.Primary, Me.trvAccess)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.ObjHak
                .GroupCode = Me.TxtGroupCode.Text
                .Description = Me.TxtDescription.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjHak

                Me.TxtPrimaryGroup.Text = .Primary
                Me.TxtGroupCode.Text = .GroupCode
                Me.TxtDescription.Text = .Description

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoTransaksi_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGroupCode.GotFocus
        Try
            If Me.TxtDescription.Enabled = True Then
                If Me.TxtGroupCode.Enabled = True Then
                    If Me.TxtGroupCode.ReadOnly = False Then
                        TxtGroupCode.BackColor = Me.BackColorAll
                    Else
                        TxtGroupCode.BackColor = Me.TxtReadonly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoTransaksi_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGroupCode.LostFocus
        Try
            If Me.TxtDescription.Enabled = True Then
                If Me.TxtGroupCode.Enabled = True Then
                    If Me.TxtGroupCode.ReadOnly = False Then
                        TxtGroupCode.BackColor = Me.TxtPrimaryGroup.BackColor
                    Else
                        TxtGroupCode.BackColor = Me.TxtReadonly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCatatan_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDescription.GotFocus
        Try
            ''If Me.TxtStockInNo.Enabled = True Then
            If Me.TxtDescription.Enabled = True Then
                If Me.TxtDescription.ReadOnly = False Then
                    Me.TxtDescription.BackColor = Me.BackColorAll
                Else
                    TxtDescription.BackColor = Me.TxtReadonly.BackColor
                End If
            End If
            ''End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCatatan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDescription.LostFocus
        Try
            ''If Me.TxtStockInNo.Enabled = True Then
            If Me.TxtDescription.Enabled = True Then
                If Me.TxtDescription.ReadOnly = False Then
                    TxtDescription.BackColor = Me.TxtPrimaryGroup.BackColor
                Else
                    TxtDescription.BackColor = Me.TxtReadonly.BackColor
                End If
            End If
            '' End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAcessDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccesDesc.GotFocus
        Try
            If Me.TxtDescription.Enabled = True Then
                If Me.TxtAccesDesc.Enabled = True Then
                    If Me.TxtAccesDesc.ReadOnly = False Then
                        TxtAccesDesc.BackColor = Me.BackColorAll
                    Else
                        TxtAccesDesc.BackColor = Me.TxtReadonly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAcessDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccesDesc.LostFocus
        Try
            If Me.TxtDescription.Enabled = True Then
                If Me.TxtAccesDesc.Enabled = True Then
                    If Me.TxtAccesDesc.ReadOnly = False Then
                        TxtAccesDesc.BackColor = Me.TxtPrimaryGroup.BackColor
                    Else
                        TxtAccesDesc.BackColor = Me.TxtReadonly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub trvAccess_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles trvAccess.AfterSelect
        Try
            If Me.TxtGroupCode.Enabled = True Then
                Me.TxtAccesDesc.Text = Me.GetDescription(e.Node)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub trvAccess_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles trvAccess.NodeMouseClick
        Try
            Me.TxtAccesDesc.Text = Me.GetDescription(e.Node)
            Me.ObjHak_Item.Set_Status_Node(Me.trvAccess)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Set_Check(ByVal ND As TreeNode, ByVal Chk As Boolean)
        Dim ND_Child As TreeNode
        Try
            ND.Checked = Chk
            For Each ND_Child In ND.Nodes
                Me.Set_Check(ND_Child, Chk)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class