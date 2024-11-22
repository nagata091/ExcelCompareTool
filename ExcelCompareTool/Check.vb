Imports System.IO
Imports Microsoft.Office.Interop

Public Class Check
    Public Function execCheck(file1Path As String, file2Path As String, gyo As Integer, retsu As Integer, lbl As Object) As String
        '初期化
        Dim result As String = ""
        Dim cntNum As Integer = 0
        Dim currentTime As Date = Now

        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkbook1 As Excel.Workbook = Nothing
        Dim xlWorkbook2 As Excel.Workbook = Nothing
        Dim resWorkBook As Excel.Workbook = Nothing
        Dim xlWorkSheet1 As Excel.Worksheet = Nothing
        Dim xlWorkSheet2 As Excel.Worksheet = Nothing
        Dim resWorkSheet As Excel.Worksheet = Nothing
        Dim tbl As Excel.ListObject = Nothing
        Dim range1Value As Object = Nothing
        Dim range2Value As Object = Nothing
        Dim range1Text As String
        Dim range2Text As String

        'Excelファイル名を取得
        Dim file1Name As String = Path.GetFileName(file1Path)
        Dim file2Name As String = Path.GetFileName(file2Path)

        'テスト結果を保存するフォルダを準備
        Dim outpurResultDir As String = Environment.CurrentDirectory() & "\テスト結果\"
        'フォルダが存在しない場合は新規で作成する
        If Not Directory.Exists(outpurResultDir) Then
            Directory.CreateDirectory(outpurResultDir)
        End If

        'テスト結果ファイルを元ファイルのコピーから生成
        Dim outputHead As String = outpurResultDir & "【テスト結果_" & Format(currentTime, "yyyymmddHHMMss") & "】"
        Dim resultFile1Path As String = outputHead & file1Name
        Dim resultFile2Path As String = outputHead & file2Name
        FileCopy(file1Path, resultFile1Path)
        FileCopy(file2Path, resultFile2Path)

        '処理結果を記録するファイル名を準備
        Dim resFilePath As String = outputHead & "Result.xlsx"

        Try
            lbl.Text = "処理中です..."
            lbl.Refresh()

            xlApp = New Excel.Application
            xlApp.Application.Visible = False
            xlApp.DisplayAlerts = False
            xlApp.ScreenUpdating = False
            xlApp.EnableEvents = False

            'ワークブックをセット
            xlWorkbook1 = xlApp.Workbooks.Open(resultFile1Path)
            xlWorkbook2 = xlApp.Workbooks.Open(resultFile2Path)
            resWorkBook = xlApp.Workbooks.Add()

            '処理結果を記録するシートの準備
            resWorkSheet = resWorkBook.Worksheets(1)
            '結果シートのヘッダーを設定
            resWorkSheet.Cells(1, 1).Value = "No"
            resWorkSheet.Cells(1, 2).Value = "シート名"
            resWorkSheet.Cells(1, 3).Value = "セル"
            resWorkSheet.Cells(1, 4).Value = "ファイル1の値"
            resWorkSheet.Cells(1, 5).Value = "ファイル2の値"
            resWorkSheet.Cells(1, 6).Value = "備考"
            '比較したファイル名を出力
            resWorkSheet.Cells(1, 8).Value = "比較元"
            resWorkSheet.Cells(1, 9).Value = file1Name
            resWorkSheet.Cells(2, 8).Value = "対象"
            resWorkSheet.Cells(2, 9).Value = file2Name

            'シートの名前と順番を比較
            If xlWorkbook1.Sheets.Count <> xlWorkbook2.Sheets.Count Then
                cntNum += 1
                resWorkSheet.Cells(cntNum + 1, 1).Value = cntNum
                resWorkSheet.Cells(cntNum + 1, 2).Value = "-"
                resWorkSheet.Cells(cntNum + 1, 3).Value = "-"
                resWorkSheet.Cells(cntNum + 1, 4).Value = xlWorkbook1.Sheets.Count
                resWorkSheet.Cells(cntNum + 1, 5).Value = xlWorkbook2.Sheets.Count
                resWorkSheet.Cells(cntNum + 1, 6).Value = "シートの数"
            End If

            For i As Integer = 1 To xlWorkbook1.Sheets.Count
                lbl.Text = "処理中です... " & i & "/" & xlWorkbook1.Sheets.Count & "番目のシートをチェック中です..."
                lbl.Refresh()

                'シートをセット
                Try
                    xlWorkSheet1 = xlWorkbook1.Sheets(i)
                Catch ex As Exception
                    xlWorkSheet1 = Nothing
                End Try
                Try
                    xlWorkSheet2 = xlWorkbook2.Sheets(i)
                Catch ex As Exception
                    xlWorkSheet2 = Nothing
                End Try

                'シート数が合わなくて、片方のシートがなくなったら処理終了
                If xlWorkSheet1 Is Nothing OrElse xlWorkSheet2 Is Nothing Then
                    cntNum += 1
                    resWorkSheet.Cells(cntNum + 1, 1).Value = cntNum
                    resWorkSheet.Cells(cntNum + 1, 2).Value = "-"
                    resWorkSheet.Cells(cntNum + 1, 3).Value = "-"
                    resWorkSheet.Cells(cntNum + 1, 4).Value = "-"
                    resWorkSheet.Cells(cntNum + 1, 5).Value = "-"
                    resWorkSheet.Cells(cntNum + 1, 6).Value = "シートなし"
                    Exit For
                End If

                'シート名のチェック。シート名が違ったら中身は見ない
                If (xlWorkSheet1.Name <> xlWorkSheet2.Name) Then
                    cntNum += 1
                    resWorkSheet.Cells(cntNum + 1, 1).Value = cntNum
                    resWorkSheet.Cells(cntNum + 1, 2).Value = xlWorkSheet1.Name
                    resWorkSheet.Cells(cntNum + 1, 3).Value = "-"
                    resWorkSheet.Cells(cntNum + 1, 4).Value = xlWorkSheet1.Name
                    resWorkSheet.Cells(cntNum + 1, 5).Value = xlWorkSheet1.Name
                    resWorkSheet.Cells(cntNum + 1, 6).Value = "シート名"
                    'シート名の背景色を変更
                    xlWorkSheet1.Tab.Color = RGB(255, 255, 0) 'イエロー
                    xlWorkSheet2.Tab.Color = RGB(255, 255, 0)
                    Continue For
                End If

                'シート内のチェック
                'シートの指定範囲を取得
                range1Value = xlWorkSheet1.Range(xlWorkSheet1.Cells(1, 1), xlWorkSheet1.Cells(gyo, retsu)).Value
                range2Value = xlWorkSheet2.Range(xlWorkSheet2.Cells(1, 1), xlWorkSheet2.Cells(gyo, retsu)).Value

                '指定列×指定行の範囲で、各セルの値が等しいかチェックする。
                For j As Integer = 1 To gyo
                    For k As Integer = 1 To retsu
                        'lbl.Text = "処理中です... " & i & "/" & xlWorkbook1.Sheets.Count & "番目のシートの(" & j & "," & k & ")セルをチェック中です..."
                        'lbl.Refresh()

                        If IsError(range1Value(j, k)) OrElse IsError(range2Value(j, k)) Then 'エラーの場合
                            cntNum += 1
                            resWorkSheet.Cells(cntNum + 1, 1).Value = cntNum
                            resWorkSheet.Cells(cntNum + 1, 2).Value = xlWorkSheet1.Name
                            resWorkSheet.Cells(cntNum + 1, 3).Value = xlWorkSheet1.Cells(j, k).Address
                            resWorkSheet.Cells(cntNum + 1, 4).Value = range1Value(j, k)
                            resWorkSheet.Cells(cntNum + 1, 5).Value = range2Value(j, k)
                            resWorkSheet.Cells(cntNum + 1, 6).Value = "エラー"
                            'セルの背景色を変更
                            xlWorkSheet1.Cells(j, k).Interior.Color = RGB(255, 255, 0)
                            xlWorkSheet2.Cells(j, k).Interior.Color = RGB(255, 255, 0)

                        ElseIf Not String.IsNullOrEmpty(range1Value(j, k)) OrElse Not String.IsNullOrEmpty(range2Value(j, k)) Then 'どちらかが空白でなければチェックする
                            'セルの値をチェック
                            range1Text = getText(range1Value(j, k))
                            range2Text = getText(range2Value(j, k))

                            If range1Text <> range2Text Then
                                cntNum += 1
                                resWorkSheet.Cells(cntNum + 1, 1).Value = cntNum
                                resWorkSheet.Cells(cntNum + 1, 2).Value = xlWorkSheet1.Name
                                resWorkSheet.Cells(cntNum + 1, 3).Value = xlWorkSheet1.Cells(j, k).Address
                                resWorkSheet.Cells(cntNum + 1, 4).Value = range1Value(j, k)
                                resWorkSheet.Cells(cntNum + 1, 5).Value = range2Value(j, k)
                                resWorkSheet.Cells(cntNum + 1, 6).Value = "値"
                                'セルの背景色を変更
                                xlWorkSheet1.Cells(j, k).Interior.Color = RGB(255, 255, 0)
                                xlWorkSheet2.Cells(j, k).Interior.Color = RGB(255, 255, 0)
                            End If
                        End If
                    Next k
                Next j
            Next i

            '結果のテーブルを作成
            tbl = resWorkSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, resWorkSheet.Range("A1").CurrentRegion, , Excel.XlYesNoGuess.xlYes)

            xlApp.ScreenUpdating = True
            xlApp.EnableEvents = True

            'Excel出力を保存
            xlWorkbook1.Save()
            xlWorkbook2.Save()
            resWorkBook.SaveAs(resFilePath)

        Catch ex As Exception
            result = ""
            Throw New Exception()
        Finally
            lbl.Text = "処理中です... 終了処理中です..."
            lbl.Refresh()

            'Excelオブジェクトをクリア
            If xlWorkbook1 IsNot Nothing Then
                xlWorkbook1.Close()
            End If
            If xlWorkbook2 IsNot Nothing Then
                xlWorkbook2.Close()
            End If
            If resWorkBook IsNot Nothing Then
                resWorkBook.Close()
            End If
            If xlApp IsNot Nothing Then
                xlApp.Quit()
            End If

            ReleaseObject(tbl)
            ReleaseObject(range1Value)
            ReleaseObject(range2Value)
            ReleaseObject(xlWorkSheet1)
            ReleaseObject(xlWorkSheet2)
            ReleaseObject(resWorkSheet)
            ReleaseObject(xlWorkbook1)
            ReleaseObject(xlWorkbook2)
            ReleaseObject(resWorkBook)
            ReleaseObject(xlApp)

            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

        result = outpurResultDir
        Return result
    End Function

    Private Function getText(value As Object) As String
        Dim text As String
        Dim strValue As String

        If Not String.IsNullOrEmpty(value) Then
            Try
                strValue = Trim(value.ToString)

                If strValue.EndsWith("%") Then
                    strValue = strValue.Replace("%", "")
                End If

                If strValue.EndsWith("円") Then
                    strValue = strValue.Replace("円", "")
                End If

                If strValue.EndsWith("p") Then
                    strValue = strValue.Replace("p", "")
                End If

                '値が数値だったら少数の細かいところまでは見ない
                If IsNumeric(strValue) Then
                    text = Left(strValue, 10)
                Else
                    text = strValue
                End If
            Catch ex As Exception
                text = ""
            End Try
        Else
            text = ""
        End If


        Return text
    End Function

    Private Sub ReleaseObject(ByVal obj As Object)
        If obj IsNot Nothing Then
            Try
                Dim intRel As Integer = 0
                Do
                    intRel = System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                Loop While intRel > 0
            Catch ex As Exception
                obj = Nothing
            Finally
                GC.Collect()
            End Try
        End If
    End Sub
End Class
