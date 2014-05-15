<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
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

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意:  以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請不要使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnActiveLoad = New System.Windows.Forms.Button()
        Me.dgvActive = New System.Windows.Forms.DataGridView()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.txtSourceXML = New System.Windows.Forms.TextBox()
        Me.btnEXPORT = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        CType(Me.dgvActive, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnActiveLoad
        '
        Me.btnActiveLoad.Location = New System.Drawing.Point(12, 12)
        Me.btnActiveLoad.Name = "btnActiveLoad"
        Me.btnActiveLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnActiveLoad.TabIndex = 0
        Me.btnActiveLoad.Text = "讀取活動"
        Me.btnActiveLoad.UseVisualStyleBackColor = True
        '
        'dgvActive
        '
        Me.dgvActive.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvActive.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvActive.Location = New System.Drawing.Point(3, 3)
        Me.dgvActive.Name = "dgvActive"
        Me.dgvActive.RowTemplate.Height = 24
        Me.dgvActive.Size = New System.Drawing.Size(417, 199)
        Me.dgvActive.TabIndex = 1
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.dgvActive, 0, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(12, 41)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(423, 205)
        Me.TableLayoutPanel1.TabIndex = 2
        '
        'txtSourceXML
        '
        Me.txtSourceXML.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSourceXML.Location = New System.Drawing.Point(94, 13)
        Me.txtSourceXML.Name = "txtSourceXML"
        Me.txtSourceXML.Size = New System.Drawing.Size(341, 22)
        Me.txtSourceXML.TabIndex = 3
        Me.txtSourceXML.Text = "http://www.tphcc.gov.tw/MainPortal/htmlcnt/rss/ActvInfo"
        '
        'btnEXPORT
        '
        Me.btnEXPORT.Location = New System.Drawing.Point(12, 252)
        Me.btnEXPORT.Name = "btnEXPORT"
        Me.btnEXPORT.Size = New System.Drawing.Size(423, 23)
        Me.btnEXPORT.TabIndex = 4
        Me.btnEXPORT.Text = "EXCEL匯出"
        Me.btnEXPORT.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(447, 286)
        Me.Controls.Add(Me.btnEXPORT)
        Me.Controls.Add(Me.txtSourceXML)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.btnActiveLoad)
        Me.Name = "frmMain"
        Me.Text = "擷取圖書館活動資訊"
        CType(Me.dgvActive, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnActiveLoad As System.Windows.Forms.Button
    Friend WithEvents dgvActive As System.Windows.Forms.DataGridView
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents txtSourceXML As System.Windows.Forms.TextBox
    Friend WithEvents btnEXPORT As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog

End Class
