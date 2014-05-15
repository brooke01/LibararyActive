Imports System.Xml
Imports BLibrary
Imports System.IO

Public Class frmMain
    Dim dtActive As DataTable

    Private Sub btnActiveLoad_Click(sender As Object, e As EventArgs) Handles btnActiveLoad.Click
        Try
            '1讀取XML來源
            Dim doc As New XmlDocument
            doc.Load(txtSourceXML.Text)

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
            dtActive = Nothing
            dtActive = dt.Clone()
            dtActive = dt.Copy
            dgvActive.DataSource = dt
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub btnEXPORT_Click(sender As Object, e As EventArgs) Handles btnEXPORT.Click
        SaveFileDialog1.FileName = "C:\匯出EXCEL.xls"
        SaveFileDialog1.InitialDirectory = "C:\"
        SaveFileDialog1.Filter = "Excel File|*.xls"
        SaveFileDialog1.Title = "請選擇EXCEL匯出儲存路徑"
        SaveFileDialog1.RestoreDirectory = True
        SaveFileDialog1.OverwritePrompt = True
        SaveFileDialog1.CheckPathExists = True
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Try
                Dim ExcelNPOI As New NPOIHelper
                Dim fs As System.IO.FileStream = New System.IO.FileStream(SaveFileDialog1.FileName, FileMode.Create, FileAccess.Write)
                Dim data As Byte() = ExcelNPOI.SaveDataToExcel(dtActive)
                fs.Write(data, 0, data.Length)
                fs.Flush()
                fs.Close() '關閉FileStream
                fs.Dispose() '釋放FileStream所有資源
                MsgBox("已匯出Excel檔", 0)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If
    End Sub
End Class
