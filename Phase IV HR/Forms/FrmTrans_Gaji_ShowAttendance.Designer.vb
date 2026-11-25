<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Gaji_ShowAttendance
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.BtnClose = New System.Windows.Forms.Button
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkshift = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkdays = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColDuration = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColAttendanceStatus = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColSchIn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchBreakOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchBreakIn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColActualIn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColActualBreakOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColActualBreakIn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColActualOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStatusRemark = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeaveCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeaveDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLeaveDoc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEarlyInDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLateDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEarlyOutDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOT = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColOTDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEarlyOTDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEarlyBreakOutDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLateBreakInDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsLate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsEarlyIn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsEarlyOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsEarlyBreakOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsLateBreakIn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchWorkDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchBreakDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColActualWorkDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColActualBreakDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsPresent = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsSick = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsLeave = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsAlfa = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsPermit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsOvertime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsEarlyOvertime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsHoliday = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryWorkShift = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLeave = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPermitDuration = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColDate, Me.ColWorkshift, Me.ColWorkdays, Me.ColDuration, Me.ColAttendanceStatus, Me.ColSchIn, Me.ColSchBreakOut, Me.ColSchBreakIn, Me.ColSchOut, Me.ColActualIn, Me.ColActualBreakOut, Me.ColActualBreakIn, Me.ColActualOut, Me.ColStatusRemark, Me.ColLeaveCode, Me.ColLeaveDesc, Me.ColLeaveDoc, Me.ColEarlyInDuration, Me.ColLateDuration, Me.ColEarlyOutDuration, Me.ColOT, Me.ColOTDuration, Me.ColEarlyOTDuration, Me.ColEarlyBreakOutDuration, Me.ColLateBreakInDuration, Me.ColRemarks, Me.ColIsLate, Me.ColIsEarlyIn, Me.ColIsEarlyOut, Me.ColIsEarlyBreakOut, Me.ColIsLateBreakIn, Me.ColSchWorkDuration, Me.ColSchBreakDuration, Me.ColActualWorkDuration, Me.ColActualBreakDuration, Me.ColIsPresent, Me.ColIsSick, Me.ColIsLeave, Me.ColIsAlfa, Me.ColIsPermit, Me.ColIsOvertime, Me.ColIsEarlyOvertime, Me.ColIsHoliday, Me.ColPrimaryWorkShift, Me.ColPrimaryLeave, Me.ColPermitDuration})
        Me.dgvItem.Location = New System.Drawing.Point(12, 12)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.Size = New System.Drawing.Size(760, 438)
        Me.dgvItem.TabIndex = 7
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(355, 220)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(75, 23)
        Me.BtnClose.TabIndex = 8
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 46
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.ReadOnly = True
        Me.ColDate.Width = 55
        '
        'ColWorkshift
        '
        Me.ColWorkshift.HeaderText = "Workshift"
        Me.ColWorkshift.Name = "ColWorkshift"
        Me.ColWorkshift.ReadOnly = True
        Me.ColWorkshift.Width = 77
        '
        'ColWorkdays
        '
        Me.ColWorkdays.HeaderText = "Workdays"
        Me.ColWorkdays.Items.AddRange(New Object() {"Workday", "Holiday"})
        Me.ColWorkdays.Name = "ColWorkdays"
        Me.ColWorkdays.ReadOnly = True
        Me.ColWorkdays.Width = 61
        '
        'ColDuration
        '
        Me.ColDuration.HeaderText = "Duration"
        Me.ColDuration.Items.AddRange(New Object() {"Full Day", "Half Day"})
        Me.ColDuration.Name = "ColDuration"
        Me.ColDuration.ReadOnly = True
        Me.ColDuration.Width = 53
        '
        'ColAttendanceStatus
        '
        Me.ColAttendanceStatus.HeaderText = "Status"
        Me.ColAttendanceStatus.Items.AddRange(New Object() {"Present", "Leave", "Half-day Leave", "Sick", "Half Day Sick", "Permit", "Half Day Permit", "Absent", "Half Absent", "Holiday", "Half Holiday"})
        Me.ColAttendanceStatus.Name = "ColAttendanceStatus"
        Me.ColAttendanceStatus.ReadOnly = True
        Me.ColAttendanceStatus.Width = 43
        '
        'ColSchIn
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchIn.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColSchIn.HeaderText = "Schedule In"
        Me.ColSchIn.Name = "ColSchIn"
        Me.ColSchIn.ReadOnly = True
        Me.ColSchIn.Width = 89
        '
        'ColSchBreakOut
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchBreakOut.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColSchBreakOut.HeaderText = "Schedule Break Out"
        Me.ColSchBreakOut.Name = "ColSchBreakOut"
        Me.ColSchBreakOut.ReadOnly = True
        Me.ColSchBreakOut.Width = 128
        '
        'ColSchBreakIn
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchBreakIn.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColSchBreakIn.HeaderText = "Schedule Break In"
        Me.ColSchBreakIn.Name = "ColSchBreakIn"
        Me.ColSchBreakIn.ReadOnly = True
        Me.ColSchBreakIn.Width = 120
        '
        'ColSchOut
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchOut.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColSchOut.HeaderText = "Schedule Out"
        Me.ColSchOut.Name = "ColSchOut"
        Me.ColSchOut.ReadOnly = True
        Me.ColSchOut.Width = 97
        '
        'ColActualIn
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColActualIn.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColActualIn.HeaderText = "Actual In"
        Me.ColActualIn.Name = "ColActualIn"
        Me.ColActualIn.ReadOnly = True
        Me.ColActualIn.Width = 74
        '
        'ColActualBreakOut
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColActualBreakOut.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColActualBreakOut.HeaderText = "Actual Break Out"
        Me.ColActualBreakOut.Name = "ColActualBreakOut"
        Me.ColActualBreakOut.ReadOnly = True
        Me.ColActualBreakOut.Width = 113
        '
        'ColActualBreakIn
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColActualBreakIn.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColActualBreakIn.HeaderText = "Actual Break In"
        Me.ColActualBreakIn.Name = "ColActualBreakIn"
        Me.ColActualBreakIn.ReadOnly = True
        Me.ColActualBreakIn.Width = 105
        '
        'ColActualOut
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColActualOut.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColActualOut.HeaderText = "Actual Out"
        Me.ColActualOut.Name = "ColActualOut"
        Me.ColActualOut.ReadOnly = True
        Me.ColActualOut.Width = 82
        '
        'ColStatusRemark
        '
        Me.ColStatusRemark.HeaderText = "Status Remarks"
        Me.ColStatusRemark.Name = "ColStatusRemark"
        Me.ColStatusRemark.ReadOnly = True
        Me.ColStatusRemark.Width = 107
        '
        'ColLeaveCode
        '
        Me.ColLeaveCode.HeaderText = "Leave Code"
        Me.ColLeaveCode.Name = "ColLeaveCode"
        Me.ColLeaveCode.ReadOnly = True
        Me.ColLeaveCode.Width = 90
        '
        'ColLeaveDesc
        '
        Me.ColLeaveDesc.HeaderText = "Leave Description"
        Me.ColLeaveDesc.Name = "ColLeaveDesc"
        Me.ColLeaveDesc.ReadOnly = True
        Me.ColLeaveDesc.Width = 118
        '
        'ColLeaveDoc
        '
        Me.ColLeaveDoc.HeaderText = "Leave Document"
        Me.ColLeaveDoc.Name = "ColLeaveDoc"
        Me.ColLeaveDoc.ReadOnly = True
        Me.ColLeaveDoc.Width = 114
        '
        'ColEarlyInDuration
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColEarlyInDuration.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColEarlyInDuration.HeaderText = "Early In"
        Me.ColEarlyInDuration.Name = "ColEarlyInDuration"
        Me.ColEarlyInDuration.ReadOnly = True
        Me.ColEarlyInDuration.Width = 67
        '
        'ColLateDuration
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColLateDuration.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColLateDuration.HeaderText = "Late"
        Me.ColLateDuration.Name = "ColLateDuration"
        Me.ColLateDuration.ReadOnly = True
        Me.ColLateDuration.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColLateDuration.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLateDuration.Width = 34
        '
        'ColEarlyOutDuration
        '
        Me.ColEarlyOutDuration.HeaderText = "Early Out"
        Me.ColEarlyOutDuration.Name = "ColEarlyOutDuration"
        Me.ColEarlyOutDuration.ReadOnly = True
        Me.ColEarlyOutDuration.Width = 75
        '
        'ColOT
        '
        Me.ColOT.HeaderText = "Overtime"
        Me.ColOT.Items.AddRange(New Object() {"No Overtime", "Weekday Overtime", "Holiday Overtime"})
        Me.ColOT.Name = "ColOT"
        Me.ColOT.ReadOnly = True
        Me.ColOT.Width = 55
        '
        'ColOTDuration
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColOTDuration.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColOTDuration.HeaderText = "OT Duration"
        Me.ColOTDuration.Name = "ColOTDuration"
        Me.ColOTDuration.ReadOnly = True
        Me.ColOTDuration.Width = 90
        '
        'ColEarlyOTDuration
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColEarlyOTDuration.DefaultCellStyle = DataGridViewCellStyle12
        Me.ColEarlyOTDuration.HeaderText = "Early OT Duration"
        Me.ColEarlyOTDuration.Name = "ColEarlyOTDuration"
        Me.ColEarlyOTDuration.ReadOnly = True
        Me.ColEarlyOTDuration.Width = 116
        '
        'ColEarlyBreakOutDuration
        '
        Me.ColEarlyBreakOutDuration.HeaderText = "Early Break Out"
        Me.ColEarlyBreakOutDuration.Name = "ColEarlyBreakOutDuration"
        Me.ColEarlyBreakOutDuration.ReadOnly = True
        Me.ColEarlyBreakOutDuration.Width = 106
        '
        'ColLateBreakInDuration
        '
        Me.ColLateBreakInDuration.HeaderText = "Late Break In"
        Me.ColLateBreakInDuration.Name = "ColLateBreakInDuration"
        Me.ColLateBreakInDuration.ReadOnly = True
        Me.ColLateBreakInDuration.Width = 96
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.ReadOnly = True
        Me.ColRemarks.Width = 74
        '
        'ColIsLate
        '
        Me.ColIsLate.HeaderText = "Is Late"
        Me.ColIsLate.Name = "ColIsLate"
        Me.ColIsLate.ReadOnly = True
        Me.ColIsLate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColIsLate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColIsLate.Visible = False
        Me.ColIsLate.Width = 45
        '
        'ColIsEarlyIn
        '
        Me.ColIsEarlyIn.HeaderText = "Is Early In"
        Me.ColIsEarlyIn.Name = "ColIsEarlyIn"
        Me.ColIsEarlyIn.ReadOnly = True
        Me.ColIsEarlyIn.Visible = False
        Me.ColIsEarlyIn.Width = 78
        '
        'ColIsEarlyOut
        '
        Me.ColIsEarlyOut.HeaderText = "Is Early Out"
        Me.ColIsEarlyOut.Name = "ColIsEarlyOut"
        Me.ColIsEarlyOut.ReadOnly = True
        Me.ColIsEarlyOut.Visible = False
        Me.ColIsEarlyOut.Width = 86
        '
        'ColIsEarlyBreakOut
        '
        Me.ColIsEarlyBreakOut.HeaderText = "Is Early Break Out"
        Me.ColIsEarlyBreakOut.Name = "ColIsEarlyBreakOut"
        Me.ColIsEarlyBreakOut.ReadOnly = True
        Me.ColIsEarlyBreakOut.Visible = False
        Me.ColIsEarlyBreakOut.Width = 117
        '
        'ColIsLateBreakIn
        '
        Me.ColIsLateBreakIn.HeaderText = "Is Late Break In"
        Me.ColIsLateBreakIn.Name = "ColIsLateBreakIn"
        Me.ColIsLateBreakIn.ReadOnly = True
        Me.ColIsLateBreakIn.Visible = False
        Me.ColIsLateBreakIn.Width = 107
        '
        'ColSchWorkDuration
        '
        Me.ColSchWorkDuration.HeaderText = "Schedule Work Duration"
        Me.ColSchWorkDuration.Name = "ColSchWorkDuration"
        Me.ColSchWorkDuration.ReadOnly = True
        Me.ColSchWorkDuration.Visible = False
        Me.ColSchWorkDuration.Width = 149
        '
        'ColSchBreakDuration
        '
        Me.ColSchBreakDuration.HeaderText = "Schedule Break Duration"
        Me.ColSchBreakDuration.Name = "ColSchBreakDuration"
        Me.ColSchBreakDuration.ReadOnly = True
        Me.ColSchBreakDuration.Visible = False
        Me.ColSchBreakDuration.Width = 151
        '
        'ColActualWorkDuration
        '
        Me.ColActualWorkDuration.HeaderText = "Actual Work Duration"
        Me.ColActualWorkDuration.Name = "ColActualWorkDuration"
        Me.ColActualWorkDuration.ReadOnly = True
        Me.ColActualWorkDuration.Visible = False
        Me.ColActualWorkDuration.Width = 134
        '
        'ColActualBreakDuration
        '
        Me.ColActualBreakDuration.HeaderText = "Actual Break Duration"
        Me.ColActualBreakDuration.Name = "ColActualBreakDuration"
        Me.ColActualBreakDuration.ReadOnly = True
        Me.ColActualBreakDuration.Width = 136
        '
        'ColIsPresent
        '
        Me.ColIsPresent.HeaderText = "Is Present"
        Me.ColIsPresent.Name = "ColIsPresent"
        Me.ColIsPresent.ReadOnly = True
        Me.ColIsPresent.Visible = False
        Me.ColIsPresent.Width = 79
        '
        'ColIsSick
        '
        Me.ColIsSick.HeaderText = "Is Sick"
        Me.ColIsSick.Name = "ColIsSick"
        Me.ColIsSick.ReadOnly = True
        Me.ColIsSick.Visible = False
        Me.ColIsSick.Width = 64
        '
        'ColIsLeave
        '
        Me.ColIsLeave.HeaderText = "Is Leave"
        Me.ColIsLeave.Name = "ColIsLeave"
        Me.ColIsLeave.ReadOnly = True
        Me.ColIsLeave.Visible = False
        Me.ColIsLeave.Width = 73
        '
        'ColIsAlfa
        '
        Me.ColIsAlfa.HeaderText = "Is Alfa"
        Me.ColIsAlfa.Name = "ColIsAlfa"
        Me.ColIsAlfa.ReadOnly = True
        Me.ColIsAlfa.Visible = False
        Me.ColIsAlfa.Width = 61
        '
        'ColIsPermit
        '
        Me.ColIsPermit.HeaderText = "Is Permit"
        Me.ColIsPermit.Name = "ColIsPermit"
        Me.ColIsPermit.ReadOnly = True
        Me.ColIsPermit.Visible = False
        Me.ColIsPermit.Width = 72
        '
        'ColIsOvertime
        '
        Me.ColIsOvertime.HeaderText = "Is Overtime"
        Me.ColIsOvertime.Name = "ColIsOvertime"
        Me.ColIsOvertime.ReadOnly = True
        Me.ColIsOvertime.Visible = False
        Me.ColIsOvertime.Width = 85
        '
        'ColIsEarlyOvertime
        '
        Me.ColIsEarlyOvertime.HeaderText = "Is Early Ovetime"
        Me.ColIsEarlyOvertime.Name = "ColIsEarlyOvertime"
        Me.ColIsEarlyOvertime.ReadOnly = True
        Me.ColIsEarlyOvertime.Visible = False
        Me.ColIsEarlyOvertime.Width = 108
        '
        'ColIsHoliday
        '
        Me.ColIsHoliday.HeaderText = "Is Holiday"
        Me.ColIsHoliday.Name = "ColIsHoliday"
        Me.ColIsHoliday.ReadOnly = True
        Me.ColIsHoliday.Visible = False
        Me.ColIsHoliday.Width = 78
        '
        'ColPrimaryWorkShift
        '
        Me.ColPrimaryWorkShift.HeaderText = "Primary Workshift"
        Me.ColPrimaryWorkShift.Name = "ColPrimaryWorkShift"
        Me.ColPrimaryWorkShift.ReadOnly = True
        Me.ColPrimaryWorkShift.Visible = False
        Me.ColPrimaryWorkShift.Width = 114
        '
        'ColPrimaryLeave
        '
        Me.ColPrimaryLeave.HeaderText = "Primary Leave"
        Me.ColPrimaryLeave.Name = "ColPrimaryLeave"
        Me.ColPrimaryLeave.ReadOnly = True
        Me.ColPrimaryLeave.Visible = False
        Me.ColPrimaryLeave.Width = 99
        '
        'ColPermitDuration
        '
        Me.ColPermitDuration.HeaderText = "Permit Duration"
        Me.ColPermitDuration.Name = "ColPermitDuration"
        Me.ColPermitDuration.ReadOnly = True
        Me.ColPermitDuration.Width = 104
        '
        'FrmTrans_Gaji_ShowAttendance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(784, 462)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.BtnClose)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(800, 500)
        Me.Name = "FrmTrans_Gaji_ShowAttendance"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Show Attendance"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkshift As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkdays As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColDuration As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColAttendanceStatus As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColSchIn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchBreakOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchBreakIn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColActualIn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColActualBreakOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColActualBreakIn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColActualOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStatusRemark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeaveCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeaveDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLeaveDoc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEarlyInDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLateDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEarlyOutDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOT As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColOTDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEarlyOTDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEarlyBreakOutDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLateBreakInDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsLate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsEarlyIn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsEarlyOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsEarlyBreakOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsLateBreakIn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchWorkDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchBreakDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColActualWorkDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColActualBreakDuration As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsPresent As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsSick As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsLeave As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsAlfa As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsPermit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsOvertime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsEarlyOvertime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsHoliday As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryWorkShift As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLeave As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPermitDuration As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
