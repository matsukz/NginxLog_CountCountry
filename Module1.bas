Sub Refresh()

    Dim DB As Worksheet: Set DB = Worksheets("Database")
    Dim Board As Worksheet: Set Board = Worksheets("Dashboard")
    
    '作業を行うワークシートをアクティブにする
    Board.Select
    
    '集計範囲のエラーをチェックする
    '集計範囲が入力されていないとき
    If Board.Range("AJ2") = "" Or Board.Range("AS2") = "" Then
        MsgBox "集計範囲が設定されていないため続行できません", vbCritical
        Exit Sub
    Else
        '何もしない
    End If
    
    'Board.Range("AJ2")のほうが小さい場合
    If Board.Range("AJ2") > Board.Range("AS2") Then
        MsgBox "集計範囲が不適切なため続行できません", vbCritical
        Exit Sub
    Else
        '何もしない
    End If
    
    '現在のデータ数をカウントする
    Dim datacount As Long: datacount = 0
    datacount = DB.ListObjects("NginxLog").ListRows.Count
    
    'MySQLに接続できるかテストする
    Select Case request("http://192.168.11.15:5500/port?ip=100.96.0.1&port=3306&option=2")
        Case 0
            'データソースを更新する
            ActiveWorkbook.RefreshAll
        Case 1
            MsgBox "MySQLサーバーへの接続試験に失敗しました。" & vbLf & "時間を空けて再試行してください。"
            Exit Sub
        Case 2
            MsgBox "Flaskからの応答がありませんでした"
            Exit Sub
        Case 3
            MsgBox "不明なエラーが発生しました"
            Exit Sub
    End Select
    
    '再クエリ後のデータをカウントする
    Dim after_datacount As Long: after_datacount = 0
    
    after_datacount = DB.ListObjects("NginxLog").ListRows.Count
    
    '円グラフ更新
    Module2.Date_Country
    
    'msgの表示
    Dim msg As String
    msg = "更新が完了しました。" & vbCrLf
    If after_datacount - datacount = 0 Then
        msg = msg & "新しいレコードはありません"
    Else
        msg = msg & after_datacount - datacount & " 件追加されました"
    End If
        
    MsgBox msg, vbInformation
    
End Sub

Function request(ByVal point As String) As Integer

    'エラーハンドリング
    On Error GoTo errorHandler
    
    'ツール　参照設定から"Microsoft XML v6.0"を有効にすること！！！
    Dim HttpReq As Object
    Set HttpReq = CreateObject("MSXML2.XMLHTTP")
    
    Dim response As Boolean: response = False
    
    'リクエスト作成 第三引数→True(非同期) False(同期)
    'キャッシュ防止の為タイムスタンプをつける(Flaskでは無視)
    Dim timestamp As String: timestamp = Format(Now, "yyyymmddhhmmss")
    HttpReq.Open "GET", point & "&nocache=" & timestamp, False
    HttpReq.send
    
    'subプロシージャに結果を渡す処理
    response = HttpReq.responseText
    If response = True Then
        request = 0 '正常な終了コード
    ElseIf response = False Then
        request = 1 'Flaskには到達したがsocket通信に失敗した
    Else
        request = 3 'その他エラー
    End If
    
    GoTo cleanUP
    
errorHandler:
    'ここを利用するときはFlaskが落ちているとき
    MsgBox "Error!! " & Err.Description & "APIサーバーをは正常ですか？", vbExclamation
    request = 3
    Set HttpReq = Nothing 'オブジェクト解放

cleanUP:
    Set HttpReq = Nothing 'オブジェクト解放
    
End Function
