<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnSelectFile1 = New System.Windows.Forms.Button()
        Me.btnSelectFile2 = New System.Windows.Forms.Button()
        Me.txtFile2 = New System.Windows.Forms.TextBox()
        Me.txtFile1 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmbGyo = New System.Windows.Forms.ComboBox()
        Me.cmbRetsu = New System.Windows.Forms.ComboBox()
        Me.btnExec = New System.Windows.Forms.Button()
        Me.lblProc = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("メイリオ", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(27, 45)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(258, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Excelファイルその１を選択　　→"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("メイリオ", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(27, 136)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(258, 24)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Excelファイルその２を選択　　→"
        '
        'btnSelectFile1
        '
        Me.btnSelectFile1.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.btnSelectFile1.Location = New System.Drawing.Point(315, 36)
        Me.btnSelectFile1.Name = "btnSelectFile1"
        Me.btnSelectFile1.Size = New System.Drawing.Size(114, 44)
        Me.btnSelectFile1.TabIndex = 4
        Me.btnSelectFile1.Text = "ファイルを選択"
        Me.btnSelectFile1.UseVisualStyleBackColor = True
        '
        'btnSelectFile2
        '
        Me.btnSelectFile2.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.btnSelectFile2.Location = New System.Drawing.Point(315, 127)
        Me.btnSelectFile2.Name = "btnSelectFile2"
        Me.btnSelectFile2.Size = New System.Drawing.Size(114, 44)
        Me.btnSelectFile2.TabIndex = 5
        Me.btnSelectFile2.Text = "ファイルを選択"
        Me.btnSelectFile2.UseVisualStyleBackColor = True
        '
        'txtFile2
        '
        Me.txtFile2.BackColor = System.Drawing.SystemColors.Control
        Me.txtFile2.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.txtFile2.Enabled = False
        Me.txtFile2.ForeColor = System.Drawing.Color.Gray
        Me.txtFile2.Location = New System.Drawing.Point(31, 177)
        Me.txtFile2.Name = "txtFile2"
        Me.txtFile2.ReadOnly = True
        Me.txtFile2.Size = New System.Drawing.Size(398, 19)
        Me.txtFile2.TabIndex = 3
        Me.txtFile2.Text = "ファイルを選択してください"
        '
        'txtFile1
        '
        Me.txtFile1.BackColor = System.Drawing.SystemColors.Control
        Me.txtFile1.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.txtFile1.Enabled = False
        Me.txtFile1.ForeColor = System.Drawing.Color.Gray
        Me.txtFile1.Location = New System.Drawing.Point(31, 86)
        Me.txtFile1.Name = "txtFile1"
        Me.txtFile1.ReadOnly = True
        Me.txtFile1.Size = New System.Drawing.Size(398, 19)
        Me.txtFile1.TabIndex = 2
        Me.txtFile1.Text = "ファイルを選択してください"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.Location = New System.Drawing.Point(28, 292)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(272, 68)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "※比較はVALUE値で行います。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "※数値は小数点以下5桁くらいまでチェックします。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "※シート名が異なる場合は比較できません。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "※比較する行数、列数が多いと結" &
    "構時間かかります。"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("メイリオ", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.Location = New System.Drawing.Point(32, 227)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(90, 24)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "行数を入力"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("メイリオ", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.Location = New System.Drawing.Point(238, 227)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(90, 24)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "列数を選択"
        '
        'cmbGyo
        '
        Me.cmbGyo.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmbGyo.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmbGyo.FormattingEnabled = True
        Me.cmbGyo.Items.AddRange(New Object() {"100", "200", "300", "400", "500", "600", "700", "800", "900", "1000", "1100", "1200", "1300", "1400", "1500", "1600", "1700", "1800", "1900", "2000"})
        Me.cmbGyo.Location = New System.Drawing.Point(128, 220)
        Me.cmbGyo.Name = "cmbGyo"
        Me.cmbGyo.Size = New System.Drawing.Size(95, 31)
        Me.cmbGyo.TabIndex = 9
        Me.cmbGyo.Text = "100"
        '
        'cmbRetsu
        '
        Me.cmbRetsu.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmbRetsu.Font = New System.Drawing.Font("メイリオ", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmbRetsu.FormattingEnabled = True
        Me.cmbRetsu.Items.AddRange(New Object() {"Z列まで", "AZ列まで", "BZ列まで", "CZ列まで"})
        Me.cmbRetsu.Location = New System.Drawing.Point(334, 220)
        Me.cmbRetsu.Name = "cmbRetsu"
        Me.cmbRetsu.Size = New System.Drawing.Size(95, 31)
        Me.cmbRetsu.TabIndex = 10
        Me.cmbRetsu.Text = "Z列まで"
        '
        'btnExec
        '
        Me.btnExec.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.btnExec.Location = New System.Drawing.Point(142, 386)
        Me.btnExec.Name = "btnExec"
        Me.btnExec.Size = New System.Drawing.Size(172, 66)
        Me.btnExec.TabIndex = 11
        Me.btnExec.Text = "実行"
        Me.btnExec.UseVisualStyleBackColor = True
        '
        'lblProc
        '
        Me.lblProc.AutoSize = True
        Me.lblProc.Font = New System.Drawing.Font("メイリオ", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblProc.ForeColor = System.Drawing.Color.DimGray
        Me.lblProc.Location = New System.Drawing.Point(12, 463)
        Me.lblProc.Name = "lblProc"
        Me.lblProc.Size = New System.Drawing.Size(80, 18)
        Me.lblProc.TabIndex = 12
        Me.lblProc.Text = "実行待機中..."
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        Me.OpenFileDialog1.Title = "Excelファイルを選択"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("メイリオ", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.WindowFrame
        Me.Label6.Location = New System.Drawing.Point(33, 66)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(217, 17)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "※チェックの比較元ファイルになります。"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(469, 490)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.lblProc)
        Me.Controls.Add(Me.btnExec)
        Me.Controls.Add(Me.cmbRetsu)
        Me.Controls.Add(Me.cmbGyo)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnSelectFile2)
        Me.Controls.Add(Me.btnSelectFile1)
        Me.Controls.Add(Me.txtFile2)
        Me.Controls.Add(Me.txtFile1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ExcelCompareTool"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents btnSelectFile1 As Button
    Friend WithEvents btnSelectFile2 As Button
    Friend WithEvents txtFile2 As TextBox
    Friend WithEvents txtFile1 As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents cmbGyo As ComboBox
    Friend WithEvents cmbRetsu As ComboBox
    Friend WithEvents btnExec As Button
    Friend WithEvents lblProc As Label
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents Label6 As Label
End Class
