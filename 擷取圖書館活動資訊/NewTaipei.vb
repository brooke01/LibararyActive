Imports System.Xml
Imports System.Web.HttpUtility

Public Class NewTaipei : Implements IReadData
    Private url As String = ""
    Private dt As DataTable

    Sub New(ByVal url As String)
        Me.url = url
    End Sub

    Public Sub Read() Implements IReadData.Read
        Try
            '1讀取XML來源
            Dim doc As New XmlDocument
            doc.Load(url)

            '2掃描XML所有項目並儲存
            Dim root As XmlElement = doc.DocumentElement '先生出一個root
            '*取得第一層所有節點
            Dim singleItemNode As XmlNode = root.SelectSingleNode("channel/item")
            For i = 0 To singleItemNode.ChildNodes.Count - 1
                dt.Columns.Add(singleItemNode.ChildNodes.Item(i).Name)
            Next
            Dim myItem As XmlNodeList = root.SelectNodes("channel/item") '取channel底下所有的item節點
            For j = 0 To myItem.Count - 1 'item個數為277
                Dim dr As DataRow = Nothing
                dr = dt.NewRow
                For k = 0 To myItem.Item(j).ChildNodes.Count - 1 'item裡的TAG數為4
                    dr(k) = HtmlDecode(myItem.Item(j).ChildNodes.Item(k).InnerXml)
                Next
                dt.Rows.Add(dr)
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub DgvDisplay(ByVal dgvActive As DataGridView)
        '3將項目顯示於dgv
        dgvActive.DataSource = dt

        For i = 0 To dgvActive.Columns.Count - 1
            dgvActive.Columns(i).Visible = False
        Next
        dgvActive.Columns("link").Visible = False

        dgvActive.Columns("title").Visible = True
        dgvActive.Columns("author").Visible = True
        dgvActive.Columns("pubDate").Visible = True
    End Sub

    Public Function getDT() As DataTable
        Return dt
    End Function
End Class
