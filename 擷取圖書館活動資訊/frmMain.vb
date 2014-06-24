Imports System.IO
Imports BLibrary

Public Class frmMain
    Dim DT_ACTIVE As DataTable

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text &= " - " & "Brooke v" & My.Application.Info.Version.ToString
        ToolStripStatusLabel1.Text = ""
        Dim dtRSSLink As New DataTable
        dtRSSLink.Columns.Add("LibName")
        dtRSSLink.Columns.Add("LibLink")

        Dim dr As DataRow
        dr = dtRSSLink.NewRow
        dr("LibName") = "新北市圖"
        dr("LibLink") = "http://www.tphcc.gov.tw/MainPortal/htmlcnt/rss/ActvInfo"
        dtRSSLink.Rows.Add(dr)

        dr = dtRSSLink.NewRow
        dr("LibName") = "台北市圖"
        dr("LibLink") = "http://www.tpml.edu.tw/search.getService.asp?serviceName=GIP.xdrss&mp=104021&ctNodeId=57441"
        dtRSSLink.Rows.Add(dr)

        dr = dtRSSLink.NewRow
        dr("LibName") = "高雄市圖"
        dr("LibLink") = "http://www.ksml.edu.tw/informactions/activity/Activity01.aspx"
        dtRSSLink.Rows.Add(dr)

        cmbRSSLink.DataSource = dtRSSLink
        cmbRSSLink.DisplayMember = "LibName"
        cmbRSSLink.ValueMember = "LibLink"

    End Sub

    Private Sub btnActiveLoad_Click(sender As Object, e As EventArgs) Handles btnActiveLoad.Click
        If cmbRSSLink.SelectedItem("libName") = "新北市圖" Then
            Dim 新北 As New NewTaipei(cmbRSSLink.SelectedValue)
            新北.Read()
            新北.DgvDisplay(dgvActive)
            DT_ACTIVE = 新北.getDT
            AddLinkColumn()
        ElseIf cmbRSSLink.SelectedItem("libName") = "台北市圖" Then
            Dim 台北 As New Taipei(cmbRSSLink.SelectedValue)
            台北.Read()
            台北.DgvDisplay(dgvActive)
            DT_ACTIVE = 台北.getDT
            AddLinkColumn()
        ElseIf cmbRSSLink.SelectedItem("libName") = "高雄市圖" Then
            Dim 高雄 As New KouSong(cmbRSSLink.SelectedValue)
            高雄.Read()
            高雄.DgvDisplay(dgvActive)
            DT_ACTIVE = 高雄.getDT
            AddLinkColumn()
        Else
            MsgBox("cmbRSSLink無此選項")
        End If
    End Sub

#Region "datagridview調整"
    Private Sub AddLinkColumn()
        Dim links As New DataGridViewLinkColumn()
        With links
            .UseColumnTextForLinkValue = True
            .HeaderText = "hyperlink"
            '.DataPropertyName = "行首文字2"
            .ActiveLinkColor = Color.White
            .LinkBehavior = LinkBehavior.SystemDefault
            .LinkColor = Color.Blue
            .TrackVisitedState = True
            .VisitedLinkColor = Color.YellowGreen
            .Text = "點選網址"
            .Name = "超連結"
        End With
        If dgvActive.Columns.Contains(links.Name) Then dgvActive.Columns.Remove(links.Name)
        dgvActive.Columns.Add(links)
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvActive.CellClick
        '當使用者點選時，會先判斷哪個column，然後再判斷點選到哪個row，最後取該cell(column)的value
        Try
            If e.ColumnIndex <> -1 Then
                Select Case dgvActive.Columns(e.ColumnIndex).Name
                    Case "超連結"
                        If e.RowIndex <> -1 Then
                            Dim url As String = dgvActive.Rows(e.RowIndex).Cells("link").Value
                            System.Diagnostics.Process.Start(url)
                            Exit Select
                        End If
                End Select
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
#End Region

#Region "EXCEL匯出"
    Private Sub btnEXPORT_Click(sender As Object, e As EventArgs) Handles btnEXPORT.Click
        SaveFileDialog1.FileName = "匯出EXCEL.xls"
        SaveFileDialog1.InitialDirectory = Directory.GetCurrentDirectory
        SaveFileDialog1.Filter = "Excel File|*.xls"
        SaveFileDialog1.Title = "請選擇EXCEL匯出儲存路徑"
        SaveFileDialog1.RestoreDirectory = True
        SaveFileDialog1.OverwritePrompt = True
        SaveFileDialog1.CheckPathExists = True
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Try
                If (DT_ACTIVE Is Nothing) = False Then
                    ExportExcelFormat(DT_ACTIVE)
                    Dim ExcelNPOI As New NPOIHelper
                    Dim fs As System.IO.FileStream = New System.IO.FileStream(SaveFileDialog1.FileName, FileMode.Create, FileAccess.Write)
                    Dim data() As Byte = Nothing
                    data = ExcelNPOI.SaveDataToExcel(DT_ACTIVE)
                    fs.Write(data, 0, data.Length)
                    fs.Flush()
                    fs.Close() '關閉FileStream
                    fs.Dispose() '釋放FileStream所有資源
                    MsgBox("完成", 0)
                Else
                    MsgBox("無資料")
                End If

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub ExportExcelFormat(DT_ACTIVE As DataTable)
        Try
            Dim list As New List(Of String)
            For Each DataColumn As DataColumn In DT_ACTIVE.Columns
                Select Case DataColumn.ColumnName
                    Case "title"
                    Case "author"
                    Case "pubDate"
                    Case Else
                        list.Add(DataColumn.ColumnName)
                End Select
            Next

            For Each s As String In list
                DT_ACTIVE.Columns.Remove(s)
            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
#End Region

End Class
