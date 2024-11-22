Imports System.IO

Public Class frmMain
    '初期化
    Dim check As New Check

    Private Sub btnExec_Click(sender As Object, e As EventArgs) Handles btnExec.Click

        '入力項目のチェック
        If txtFile1.Text = "ファイルを選択してください" OrElse txtFile2.Text = "ファイルを選択してください" Then
            MsgBox("チェック対象が未選択です。")
            Exit Sub
        End If

        '入力項目の取得
        Dim file1Path As String = txtFile1.Text
        Dim file2Path As String = txtFile2.Text
        Dim gyo As Integer
        If cmbGyo.Text <> "" Then
            gyo = Integer.Parse(cmbGyo.Text)
        Else
            MsgBox("チェックする行数を入力してください。")
            Exit Sub
        End If
        Dim retsu As Integer
        If cmbRetsu.Text = "Z列まで" Then
            retsu = 26
        ElseIf cmbRetsu.Text = "AZ列まで" Then
            retsu = 26 * 2
        ElseIf cmbRetsu.Text = "BZ列まで" Then
            retsu = 26 * 3
        ElseIf cmbRetsu.Text = "CZ列まで" Then
            retsu = 26 * 4
        Else
            MsgBox("対象列をリストから選択してください。")
            Exit Sub
        End If

        'チェック対象のファイル名を取得
        Dim file1Name As String = Path.GetFileName(file1Path)
        Dim file2Name As String = Path.GetFileName(file2Path)

        '実行確認
        Dim execMessage As String
        execMessage = file1Name & vbCrLf & file2Name & vbCrLf & "以上のファイルを" & vbCrLf & gyo.ToString & "行 × " & cmbRetsu.SelectedItem & vbCrLf & "チェックします。開始しますか？"
        Dim execResult As DialogResult = MessageBox.Show(execMessage, "実行確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If execResult = DialogResult.No Then
            Exit Sub
        End If

        Me.Enabled = False
        Me.UseWaitCursor = True

        lblProc.Text = "処理を開始中です..."
        lblProc.Refresh()

        'チェック処理を実行
        Dim res As String = ""
        Try
            res = check.execCheck(file1Path, file2Path, gyo, retsu, lblProc)
        Catch ex As Exception
            Me.Enabled = True
            Me.UseWaitCursor = False
            res = ""
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

        If res <> "" Then
            Me.Enabled = True
            Me.UseWaitCursor = False
            lblProc.Text = "完了しました！"
            lblProc.Refresh()

            '結果を保存したフォルダを開く
            Process.Start("explorer.exe", res)
        Else
            Me.Enabled = True
            Me.UseWaitCursor = False
            lblProc.Text = "失敗しました。"
            lblProc.Refresh()
        End If

    End Sub

    Private Sub btnSelectFile1_Click(sender As Object, e As EventArgs) Handles btnSelectFile1.Click
        If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then
            txtFile1.Text = OpenFileDialog1.FileName

            lblProc.Text = "実行待機中..."
            lblProc.Refresh()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub btnSelectFile2_Click(sender As Object, e As EventArgs) Handles btnSelectFile2.Click
        If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then
            txtFile2.Text = OpenFileDialog1.FileName

            lblProc.Text = "実行待機中..."
            lblProc.Refresh()
        Else
            Exit Sub
        End If
    End Sub
End Class
