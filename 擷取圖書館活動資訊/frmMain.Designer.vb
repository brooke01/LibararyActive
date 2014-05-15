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
        Me.btnEXPORT = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.cmbRSSLink = New System.Windows.Forms.ComboBox()
        CType(Me.dgvActive, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnActiveLoad
        '
        Me.btnActiveLoad.Location = New System.Drawing.Point(3, 3)
        Me.btnActiveLoad.Name = "btnActiveLoad"
        Me.btnActiveLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnActiveLoad.TabIndex = 0
        Me.btnActiveLoad.Text = "讀取活動"
        Me.btnActiveLoad.UseVisualStyleBackColor = True
        '
        'dgvActive
        '
        Me.dgvActive.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.TableLayoutPanel2.SetColumnSpan(Me.dgvActive, 2)
        Me.dgvActive.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvActive.Location = New System.Drawing.Point(3, 57)
        Me.dgvActive.Name = "dgvActive"
        Me.dgvActive.RowTemplate.Height = 24
        Me.dgvActive.Size = New System.Drawing.Size(407, 106)
        Me.dgvActive.TabIndex = 1
        '
        'btnEXPORT
        '
        Me.TableLayoutPanel2.SetColumnSpan(Me.btnEXPORT, 2)
        Me.btnEXPORT.Location = New System.Drawing.Point(3, 169)
        Me.btnEXPORT.Name = "btnEXPORT"
        Me.btnEXPORT.Size = New System.Drawing.Size(404, 23)
        Me.btnEXPORT.TabIndex = 4
        Me.btnEXPORT.Text = "EXCEL匯出"
        Me.btnEXPORT.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 2
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.cmbRSSLink, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.dgvActive, 0, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.btnEXPORT, 0, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.btnActiveLoad, 0, 0)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 3
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 32.78008!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 67.21992!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 95.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(422, 286)
        Me.TableLayoutPanel2.TabIndex = 5
        '
        'cmbRSSLink
        '
        Me.cmbRSSLink.FormattingEnabled = True
        Me.cmbRSSLink.Location = New System.Drawing.Point(220, 8)
        Me.cmbRSSLink.Name = "cmbRSSLink"
        Me.cmbRSSLink.Size = New System.Drawing.Size(187, 20)
        Me.cmbRSSLink.TabIndex = 6
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(422, 286)
        Me.Controls.Add(Me.TableLayoutPanel2)
        Me.Name = "frmMain"
        Me.Text = "擷取圖書館活動資訊"
        CType(Me.dgvActive, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnActiveLoad As System.Windows.Forms.Button
    Friend WithEvents dgvActive As System.Windows.Forms.DataGridView
    Friend WithEvents btnEXPORT As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents cmbRSSLink As System.Windows.Forms.ComboBox

End Class
