Imports System.Xml
Imports BLibrary
Imports System.IO
Imports System.Text

Public Class frmMain
    Dim DT_ACTIVE As DataTable

    Private Sub btnActiveLoad_Click(sender As Object, e As EventArgs) Handles btnActiveLoad.Click
        If cmbRSSLink.SelectedText = "新北市圖" Or cmbRSSLink.SelectedText = "台北市圖" Then
            readXmlMethod()
        ElseIf cmbRSSLink.SelectedText = "高雄市圖" Then
            readHtmlMethod()
        Else
            MsgBox("cmbRSSLink無此選項")
        End If
    End Sub

    Private Sub readHtmlMethod()
        Dim webBrower As New WebBrowser
        webBrower.Navigate(cmbRSSLink.SelectedValue)
        Dim Htmldoc As HtmlDocument

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

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text &= " - " & "創作人:Brooke v" & My.Application.Info.Version.ToString
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
        dr("LibLink") = "http://www.ksml.edu.tw/informactions/activity/Activity01.aspx?Page=1"
        dtRSSLink.Rows.Add(dr)

        cmbRSSLink.DataSource = dtRSSLink
        cmbRSSLink.DisplayMember = "LibName"
        cmbRSSLink.ValueMember = "LibLink"

    End Sub



End Class
