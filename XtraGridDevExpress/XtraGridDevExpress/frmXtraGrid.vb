#Region "ABOUT"
' / --------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gsoft.com/
' /
' / Purpose: Sample code to used XtraGrid DevExpress.
' / Microsoft Visual Basic .NET (2010) + MS Access 2007+
' /
' / This is open source code under @Copyleft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / --------------------------------------------------------------------
#End Region

Imports System.Data.OleDb
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.Data
Imports DevExpress.XtraGrid

Public Class frmXtraGrid
    '// หากเป็นโปรเจคจริงๆ กลุ่มตัวแปรเหล่านี้ต้องนำไปวางไว้ใน Module 
    Dim Conn As OleDb.OleDbConnection
    Dim DA As New System.Data.OleDb.OleDbDataAdapter()
    Dim Cmd As New System.Data.OleDb.OleDbCommand
    Dim DT As New DataTable
    Dim strSQL As String
    '// Connect MS Access DataBase
    Function ConnectDataBase() As OleDb.OleDbConnection
        Return New OleDb.OleDbConnection( _
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
            MyPath(Application.StartupPath) & "data\" & "Countries.accdb;Persist Security Info=True")
    End Function

    ' / --------------------------------------------------------------------------------
    ' / Get my project path
    Function MyPath(ByVal AppPath As String) As String
        '/ MessageBox.Show(AppPath);
        MyPath = AppPath.ToLower.Replace("\bin\debug", "\").Replace("\bin\release", "\").Replace("\bin\x86\debug", "\")
        '/ Return Value
        '// If not found folder then put the \ (BackSlash) at the end.
        If Microsoft.VisualBasic.Right(MyPath, 1) <> Chr(92) Then MyPath = MyPath & Chr(92)
    End Function

    '// Start Here.
    Private Sub frmXtraGrid_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        DevExpress.Utils.AppearanceObject.DefaultFont = New Font("Tahoma", 10, FontStyle.Regular, GraphicsUnit.Point, 0)
        Conn = ConnectDataBase()
        '// Anchor @Runtime
        Me.GridControl1.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
        With cmbDisplay
            .Items.Add("แสดงผลทั้งหมด")
            .Items.Add("จัดกลุ่มตาม Zone")
        End With
        cmbDisplay.SelectedIndex = 0
        '// จะกระโดดไปเหตุการณ์ SelectedIndexChanged ของ ComboBox
    End Sub

    Private Sub cmbDisplay_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbDisplay.SelectedIndexChanged
        Select Case cmbDisplay.SelectedIndex
            Case 0
                '// Refresh Data.
                Call GetDataTable(False)
                Call SetupGridView()
                Call GroupSumFooter()

            Case 1
                Call GroupSumFooter()
                Call GroupByZone()
        End Select
    End Sub

    Private Sub SetupGridView()
        GridView1.Columns.Clear()
        '// ตั้งค่าคุณสมบัติ XtraGrid
        '// Start Add Fields.
        Dim GC As New GridColumn    '// Imports DevExpress.XtraGrid.Columns
        GC = GridView1.Columns.AddField("CountryPK")
        With GC
            .Caption = "CountryPK"
            .UnboundType = DevExpress.Data.UnboundColumnType.Integer
            '// ซ่อนหลัก Index = 0 ซึ่งเป็นค่า Primary Key 
            '// เมื่อผู้ใช้กดดับเบิ้ลคลิ๊กเมาส์ หรือกด Enter ในแต่ละแถว เราจะนำค่านี้ไป Query เพื่อแสดงผลรายละเอียดอีกที
            .Visible = False
        End With
        '// รูปภาพแบบ BLOB (Binary Large OBject) แสดงตัวอย่างในการนำภาพมาแสดงผล
        GC = GridView1.Columns.AddField("Flag")
        With GC
            .Caption = "Flag"
            .UnboundType = DevExpress.Data.UnboundColumnType.Object
            .AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            .AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            .Visible = True
            .Width = 30
        End With
        '//
        GC = GridView1.Columns.AddField("A2")
        With GC
            .Caption = "A2"
            .UnboundType = DevExpress.Data.UnboundColumnType.String
            .Visible = True
        End With
        '//
        GC = GridView1.Columns.AddField("Country")
        With GC
            .Caption = "Country"
            .UnboundType = DevExpress.Data.UnboundColumnType.String
            .Visible = True
        End With
        '
        GC = GridView1.Columns.AddField("Capital")
        With GC
            .Caption = "Capital"
            .UnboundType = DevExpress.Data.UnboundColumnType.String
            .Visible = True
        End With
        '//
        GC = GridView1.Columns.AddField("ZoneName")
        With GC
            .Caption = "ZoneName"
            .UnboundType = DevExpress.Data.UnboundColumnType.String
            '// หากมีการจัดกลุ่มต้องปิดการแสดงผลหลักนี้
            If cmbDisplay.SelectedIndex = 0 Then
                .Visible = True
            Else
                .Visible = False
            End If
        End With
        '//
        GC = GridView1.Columns.AddField("Population")
        With GC
            .Caption = "Population"
            .UnboundType = DevExpress.Data.UnboundColumnType.Decimal
            .DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
            .DisplayFormat.FormatString = "#,##;(-#,##.00);0"
            .AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
            .AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
            .Visible = True
        End With

        '// Setup XtraGrid Properties @Run Time.
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView = CType(GridControl1.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
        DevExpress.XtraEditors.WindowsFormsSettings.DefaultFont = New Font("Tahoma", 10)
        With view
            .OptionsBehavior.AutoPopulateColumns = False
            '// ไม่ให้เกิด GroupPanel เพื่อจับหลักลากมาวางในส่วนนี้ได้ จะใช้การเขียนโค้ดในการจัดกลุ่มแทน
            .OptionsView.ShowGroupPanel = False
            .OptionsCustomization.AllowFilter = True '//False
            .OptionsCustomization.AllowColumnMoving = False
            '// CheckBox Selector
            .OptionsSelection.ResetSelectionClickOutsideCheckboxSelector = True
        End With
        '//
        With GridView1
            .RowHeight = 25
            .GroupRowHeight = 30
            .Appearance.Row.Font = New Font("Tahoma", 10, FontStyle.Regular)
            .Appearance.HeaderPanel.Font = New Font("Tahoma", 10, FontStyle.Bold)
            .Appearance.GroupRow.Font = New Font("Tahoma", 10, FontStyle.Bold)
            .Appearance.FocusedRow.Font = New Font("Tahoma", 10, FontStyle.Bold)
            .Appearance.SelectedRow.Font = New Font("Tahoma", 10)
            .Appearance.FooterPanel.Font = New Font("Tahoma", 10)
            ' / Word wrap
            .OptionsView.RowAutoHeight = True
            .Appearance.Row.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap

            '// ปรับการเพิ่ม/แก้ไข/ลบข้อมูล
            .OptionsBehavior.AllowAddRows = False
            .OptionsBehavior.AllowDeleteRows = False
            .OptionsBehavior.Editable = False '// True
            '// Display
            .FocusRectStyle = DrawFocusRectStyle.RowFocus
            .Appearance.FocusedCell.ForeColor = Color.Red
            .Appearance.FocusedCell.Options.UseTextOptions = True
            .OptionsSelection.EnableAppearanceFocusedCell = False '//True
            '// Make the group footers always visible.
            .OptionsView.ShowFooter = True ' False
            .OptionsView.GroupFooterShowMode = GroupFooterShowMode.VisibleAlways
        End With
    End Sub

    '// การค้นหาข้อมูล หรือแสดงผลข้อมูลทั้งหมด จะใช้เพียงโปรแกรมย่อยแบบ Sub Program เดียวกัน
    '// หากค่าที่ส่งมาเป็น False (หรือไม่ส่งมา จะถือเป็นค่า Default) นั่นคือให้แสดงผลข้อมูลทั้งหมด
    '// หากค่าที่ส่งมาเป็น True จะเป็นการค้นหาข้อมูล
    Private Sub GetDataTable(Optional ByVal blnSearch As Boolean = False)
        strSQL = _
            " SELECT CountryPK, Flag, A2, Country, Capital, Population, Zones.ZoneName " & _
            " FROM Countries INNER JOIN Zones ON Countries.ZoneFK = Zones.ZonePK "
        '// เป็นการค้นหา
        If blnSearch Then
            strSQL = strSQL & _
                    " WHERE " & _
                    " [A2] " & " Like '%" & Trim(txtSearch.Text) & "%'" & " OR " & _
                    " [Country] " & " Like '%" & Trim(txtSearch.Text) & "%'" & " OR " & _
                    " [Capital] " & " Like '%" & Trim(txtSearch.Text) & "%'" & " OR " & _
                    " [ZoneName] " & " Like '%" & Trim(txtSearch.Text) & "%'"
            '// Else ไม่ต้องมี
        End If
        '// เอา strSQL มาเรียงต่อกัน
        strSQL = strSQL & " ORDER BY Countries.A2 "
        '/
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        Try
            Cmd = Conn.CreateCommand
            Cmd.CommandText = strSQL
            DT = New DataTable
            DA = New OleDbDataAdapter(Cmd)
            DA.Fill(DT)
            '// Bound Data
            GridControl1.DataSource = DT
            lblRecordCount.Text = "[จำนวน : " & DT.Rows.Count & " รายการ.]"
            DT.Dispose()
            DA.Dispose()
            txtSearch.Text = ""
            txtSearch.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    '// Footer for show Sum Total of Population.
    Sub GroupSumFooter()
        GridView1.ClearGrouping()
        GridView1.GroupSummary.Clear()
        '/ Make sure the group footers always visible.
        GridView1.OptionsView.ShowFooter = True
        GridView1.OptionsView.GroupFooterShowMode = GroupFooterShowMode.VisibleAlways
        '// Creating sum group summaries.
        Dim SumPopulation As New GridColumn
        SumPopulation = GridView1.Columns("Population")
        SumPopulation.SummaryItem.SummaryType = SummaryItemType.Sum
        '// format the summary of population.
        SumPopulation.SummaryItem.DisplayFormat = "รวมทั้งหมด : {0:N0}"
        '// group summary
        GridView1.GroupSummary.Add(SummaryItemType.Sum, "Population")
    End Sub

    '// จัดกลุ่มตาม Zone โดยใช้การเขียนโค้ดแทนการจับหลักลากมาวาง
    Private Sub GroupByZone()
        GridView1.ClearGrouping()
        GridView1.GroupSummary.Clear()
        '/ Import DevExpress.XtraGrid.Views.Grid
        GridView1.Columns("ZoneName").GroupIndex = 1
        GridView1.OptionsView.GroupFooterShowMode = GroupFooterShowMode.VisibleIfExpanded
        '/ Create and setup the Group with ZoneName.
        Dim ZoneItem As GridGroupSummaryItem = New GridGroupSummaryItem()
        ZoneItem.FieldName = "ZoneName"
        '// นับจำนวนประเทศแสดงผลใน Group Header ของแต่ละกลุ่ม
        ZoneItem.DisplayFormat = "| จำนวน : {0:N0} ประเทศ"
        ZoneItem.SummaryType = DevExpress.Data.SummaryItemType.Count
        GridView1.GroupSummary.Add(ZoneItem)

        '/ Create and setup the summary item.
        Dim SumPopulation As GridGroupSummaryItem = New GridGroupSummaryItem()
        SumPopulation.FieldName = "Population"
        SumPopulation.SummaryType = DevExpress.Data.SummaryItemType.Sum
        SumPopulation.DisplayFormat = "รวม : {0:N0}"
        SumPopulation.ShowInGroupColumnFooter = GridView1.Columns("Population")
        GridView1.GroupSummary.Add(SumPopulation)
    End Sub

    '// แสดงผลทั้งหมด
    Private Sub btnRefresh_Click(sender As System.Object, e As System.EventArgs) Handles btnRefresh.Click
        '// Refresh Data.
        Call GetDataTable()
        Call SetupGridView()
        Call GroupSumFooter()
    End Sub

    '// การค้นหา
    Private Sub txtSearch_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearch.KeyPress
        If Trim(txtSearch.Text.Trim) = "" Or Len(Trim(txtSearch.Text.Trim)) = 0 Then Return
        '// RetrieveData(True) It means searching for information.
        If e.KeyChar = Chr(13) Then '// Press Enter
            '// No beep.
            e.Handled = True
            '// Undesirable characters for the database ex.  ', * or %
            txtSearch.Text = txtSearch.Text.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
            Call SetupGridView()
            If cmbDisplay.SelectedIndex = 0 Then
                Call GetDataTable(True)
                Call GroupSumFooter()
            Else
                Call GetDataTable(True)
                Call GroupSumFooter()
                Call GroupByZone()
            End If
        End If
    End Sub

    '// เหตุการณ์ดับเบิ้ลคลิ๊กเมาส์ในแต่ละแถวของตารางกริด
    Private Sub GridControl1_DoubleClick(sender As Object, e As System.EventArgs) Handles GridControl1.DoubleClick
        If TryCast(GridControl1.FocusedView, GridView).RowCount = 0 Then Exit Sub
        '// การรับค่าในแต่ละหลักของตารางกริด
        Dim PK As Integer = CLng(GridView1.GetRowCellValue(GridView1.GetSelectedRows(0), "CountryPK"))
        Dim Country As String = GridView1.GetRowCellValue(GridView1.GetSelectedRows(0), "Country")
        Dim Capital As String = GridView1.GetRowCellValue(GridView1.GetSelectedRows(0), "Capital")
        'MsgBox(PK)
        '// หาก PK = 0 จะเป็น Group Header
        If PK <> 0 Then DevExpress.XtraEditors.XtraMessageBox.Show( _
            "Primary Key is : " & PK & vbCrLf & _
            "Country is: " & Country & vbCrLf & _
            "Capital is: " & Capital _
            )
    End Sub

    '// เหตุการณ์กด Enter ในแต่ละแถวของตารางกริด
    Private Sub GridControl1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles GridControl1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call GridControl1_DoubleClick(sender, e)
            '// ป้องกันไม่ให้เลื่อนโฟกัสไปแถวถัดไป
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub frmXtraGrid_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If Conn.State = ConnectionState.Open Then Conn.Close()
        Me.Dispose()
        GC.SuppressFinalize(Me)
        Application.Exit()
    End Sub

End Class
