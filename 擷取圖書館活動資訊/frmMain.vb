Imports System.Xml
Imports BLibrary
Imports System.IO
Imports System.Text
Imports HtmlAgilityPack
Imports System.Threading
Imports System.Text.RegularExpressions

Public Class frmMain
    Dim DT_ACTIVE As DataTable
    Dim webBrower As New WebBrowser
    Dim loading As Boolean = True
    Dim 高雄tatalPage As Integer = -1
    Dim 高雄htmlStream As Stream

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text &= " - " & "創作人:Brooke v" & My.Application.Info.Version.ToString
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
        If cmbRSSLink.SelectedItem("libName") = "新北市圖" Or cmbRSSLink.SelectedItem("libName") = "台北市圖" Then
            readXmlMethod()
        ElseIf cmbRSSLink.SelectedItem("libName") = "高雄市圖" Then
            readHtmlMethod()
        Else
            MsgBox("cmbRSSLink無此選項")
        End If
    End Sub

    Private Sub readXmlMethod()
        Try
            '1讀取XML來源
            Dim doc As New XmlDocument
            doc.Load(cmbRSSLink.SelectedValue)

            '2掃描XML所有項目並儲存
            Dim root As XmlElement = doc.DocumentElement '先生出一個root
            '*取得第一層所有節點
            Dim singleItemNode As XmlNode = root.SelectSingleNode("channel/item")
            Dim dt As New DataTable
            For i = 0 To singleItemNode.ChildNodes.Count - 1
                dt.Columns.Add(singleItemNode.ChildNodes.Item(i).Name)
            Next
            Dim nodes As XmlNodeList = root.SelectNodes("channel/item") '取channel底下所有的item節點
            For j = 0 To nodes.Count - 1 'item個數為277
                Dim dr As DataRow = Nothing
                dr = dt.NewRow
                For k = 0 To nodes.Item(j).ChildNodes.Count - 1 'item裡的TAG數為4
                    dr(k) = nodes.Item(j).ChildNodes.Item(k).InnerXml
                Next
                dt.Rows.Add(dr)
            Next

            '3將項目顯示於dgv
            DT_ACTIVE = Nothing
            DT_ACTIVE = dt.Clone()
            DT_ACTIVE = dt.Copy
            dgvActive.DataSource = dt
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

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

#Region "高雄特殊讀取方式"
    Private Sub readHtmlMethod()
        Try
            AddHandler webBrower.Navigated, AddressOf WebBrowser1_Navigated '指定webBrower的Navigated事件給WebBrowser1_Navigated方法處理
            AddHandler Timer1.Tick, AddressOf GetHtml
            Timer1.Interval = 3000

            Dim dt As New DataTable
            dt.Columns.Add("title")
            dt.Columns.Add("link")
            dt.Columns.Add("author")
            dt.Columns.Add("pubDate")

            '先決定活動頁面有幾頁
            loading = True
            webBrower.Navigate(cmbRSSLink.SelectedValue) '在Navigate時loading指標為true
            '在迴圈內是靠timer計數器來讀取網頁，讀完,loading並設成false
            While (loading)
                Application.DoEvents()
            End While
            ToolStripStatusLabel1.Text = ""

            For i = 1 To 高雄tatalPage
                loading = True
                webBrower.Navigate(cmbRSSLink.SelectedValue & "?Page=" & i.ToString & "")
                While (loading)
                    Application.DoEvents()
                End While
                ConvwertDataTable(dt)
            Next
            ToolStripStatusLabel1.Text = ""

            'to datagridview
            DT_ACTIVE = dt.Clone
            DT_ACTIVE = dt.Copy
            dgvActive.DataSource = DT_ACTIVE

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub WebBrowser1_Navigated(sender As Object, e As WebBrowserNavigatedEventArgs)
        Timer1.Start()
        ToolStripStatusLabel1.Text = "讀取中"
    End Sub

    Private Sub ConvwertDataTable(ByRef dt As DataTable)
        Try
            Dim Htmldoc As New HtmlDocument
            Htmldoc.Load(高雄htmlStream, Encoding.UTF8)

            Dim HtmlNodeCollection As HtmlNodeCollection = Htmldoc.DocumentNode.SelectNodes("//table[@id='ctl00_ContentPlaceHolder1_DataList1']/tr/td/table/tr/td/table/tr")
            HtmlNodeCollection.RemoveAt(0) '移除標題

            For Each HtmlNode As HtmlNode In HtmlNodeCollection
                '範例字串
                '            <tr>
                '                            <td width="2%"><img alt="*" src="../Images/icon03_001.gif" width="12" height="13" align="absmiddle"></td>                                
                '                            <td width="52%"><a href="activity_01.aspx?code=0000002797">閱讀起步走－「幼兒園暨托兒所嬰幼兒閱讀推廣研習」錄取名單(幼兒園及托兒所部份)</a></td>
                '                            <td width="16%" align="center"><strong>市圖總館</strong></td>
                '                            <td width="15%" class="f05" align="center">[2014-05-09]</td>
                '                            <td width="15%" class="f05" align="center">

                '                           &nbsp;

                '</td>                                
                '                            </tr>

                '取得指定內容
                Dim dr As DataRow
                dr = dt.NewRow
                dr("title") = Regex.Match(HtmlNode.OuterHtml, ".*<a href.*>(.*?)</a>").Groups(1).Value
                dr("link") = ""
                dr("author") = Regex.Match(HtmlNode.OuterHtml, ".*<strong.*>(.*?)</strong>").Groups(1).Value
                dr("pubDate") = Regex.Match(HtmlNode.OuterHtml, ".*align=""center"">(\[.*?)</td>").Groups(1).Value
                dt.Rows.Add(dr)
            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub GetHtml(sender As Object, e As EventArgs)
        Timer1.Stop()
        If 高雄tatalPage = -1 Then
            高雄tatalPage = 1 'webBrower.DocumentText
        Else
            高雄htmlStream = webBrower.DocumentStream
        End If
        loading = False
    End Sub

#End Region

End Class
