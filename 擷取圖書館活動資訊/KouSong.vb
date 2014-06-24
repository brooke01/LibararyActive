Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports HtmlAgilityPack

Public Class KouSong : Implements IReadData
    Dim webBrower As New WebBrowser
    Dim loading As Boolean = True
    Dim 高雄tatalPage As Integer = -1
    Dim 高雄htmlStream As Stream
    Dim Timer1 As New Timer

    Dim ToolStripStatusLabel1 As ToolStripStatusLabel
    Private url As String = ""
    Private dt As New DataTable

    Sub New(ByVal url As String)
        Me.url = url
    End Sub

    Public Sub setToolStripStatusLabel(ByValToolStripStatusLabel1 As ToolStripStatusLabel)
        Me.ToolStripStatusLabel1 = ToolStripStatusLabel1
    End Sub

    Public Sub Read() Implements IReadData.Read
        Try
            AddHandler webBrower.Navigated, AddressOf WebBrowser1_Navigated '指定webBrower的Navigated事件給WebBrowser1_Navigated方法處理
            AddHandler Timer1.Tick, AddressOf GetHtml
            Timer1.Interval = 2000

            dt.Columns.Add("title")
            dt.Columns.Add("link")
            dt.Columns.Add("author")
            dt.Columns.Add("pubDate")

            '先決定活動頁面有幾頁
            loading = True
            webBrower.Navigate(url) '在Navigate時loading指標為true
            '在迴圈內是靠timer計數器來讀取網頁，讀完,loading並設成false
            While (loading)
                Application.DoEvents()
            End While
            ToolStripStatusLabel1.Text = ""

            For i = 1 To 高雄tatalPage
                loading = True
                webBrower.Navigate(url & "?Page=" & i.ToString & "")
                While (loading)
                    Application.DoEvents()
                End While
                ConvwertDataTable(dt)
            Next
            ToolStripStatusLabel1.Text = ""

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
                dr("link") = "http://www.ksml.edu.tw/informactions/activity/" & Regex.Match(HtmlNode.OuterHtml, ".*<a href=""(.*?)"">").Groups(1).Value
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
            高雄tatalPage = Get高雄tatalPage()
        Else
            高雄htmlStream = webBrower.DocumentStream
        End If
        loading = False
    End Sub

    Private Function Get高雄tatalPage() As Integer
        Dim s As String = webBrower.DocumentText
        s = Regex.Match(s, ".*<span id=""ctl00_ContentPlaceHolder1_lblCurPage"">第1頁,共(.*?)筆</span>").Groups(1).Value
        Return CInt(s) / 10 + 1
    End Function

    Public Sub DgvDisplay(ByVal dgvActive As DataGridView)
        '3將項目顯示於dgv
        dgvActive.DataSource = dt
         dgvActive.Columns("link").Visible = False
    End Sub

    Public Function getDT() As DataTable
        Return dt
    End Function
End Class
