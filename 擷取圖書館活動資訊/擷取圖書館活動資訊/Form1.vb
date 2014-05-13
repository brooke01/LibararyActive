Imports System.Xml

Public Class Form1

    Private Sub btnActiveLoad_Click(sender As Object, e As EventArgs) Handles btnActiveLoad.Click
        Try
            '1讀取XML來源
            Dim doc As New XmlDocument
            doc.Load(txtSourceXML.Text)

            '2掃描XML所有項目並儲存
            Dim root As XmlElement = doc.DocumentElement '先生出一個root
            '*取得第一層所有節點
            Dim singleItemNode As XmlNode = root.SelectSingleNode("channel/item")
            '3將項目顯示於dgv
            Dim dt As New DataTable
            For i = 0 To singleItemNode.ChildNodes.Count - 1
                dt.Columns.Add(singleItemNode.ChildNodes.Item(i).Name)
            Next

            '版本一
            Dim nodes As XmlNode = root.SelectSingleNode("channel") '取channel底下所有的節點包括非item '左括號就算是一個Node
            For j = 0 To nodes.ChildNodes.Count - 1 'item個數為277
                Dim dr As DataRow
                dr = dt.NewRow 'item裡的TAG數為4
                For k = 0 To nodes.ChildNodes.Item(j).ChildNodes.Count - 1
                    dr(k) = nodes.ChildNodes.Item(j).ChildNodes.Item(k).InnerText
                Next
                dt.Rows.Add(dr)
            Next

            '版本二
            'Dim nodes As XmlNodeList = root.SelectNodes("channel/item") '取channel底下所有的item節點
            'For j = 0 To nodes.Count - 1 'item個數為277
            '    Dim dr As DataRow = Nothing
            '    dr = dt.NewRow
            '    For k = 0 To nodes.Item(j).ChildNodes.Count - 1 'item裡的TAG數為4
            '        dr(k) = nodes.Item(j).ChildNodes.Item(k).InnerXml
            '    Next
            '    dt.Rows.Add(dr)
            'Next

            dgvActive.DataSource = dt
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
End Class
