Attribute VB_Name = "Module1"
Sub Refresh()
    
    '作業を行うワークシートをアクティブにする
    Worksheets("Database").Select
    
    '現在のデータ数をカウントする
    Dim datacount As Long
    datacount = 0
    
    datacount = Worksheets("Database").ListObjects("NginxLog").ListRows.Count

    'MySQLに接続できるかテストする
    Select Case request("http://192.168.11.15:5500/port?ip=100.96.0.1&port=3306&option=2")
        Case 0
            'データソースを更新する
            ActiveWorkbook.RefreshAll
        Case 1
            MsgBox "MySQLサーバーへの接続試験に失敗しました。" & vbLf & "時間を空けて再試行してください。"
            Worksheets("Dashboard").Select
            Exit Sub
        Case 2
            MsgBox "Flaskからの応答がありませんでした"
            Worksheets("Dashboard").Select
            Exit Sub
        Case 3
            MsgBox "不明なエラーが発生しました"
            Worksheets("Dashboard").Select
            Exit Sub
    End Select
        
    
    '再照会後のデータをカウントする
    Dim after_datacount As Long
    after_datacount = 0
    
    after_datacount = Worksheets("Database").ListObjects("NginxLog").ListRows.Count
    
    Debug.Print (datacount)
    Debug.Print (after_datacount)
    
    '円グラフ更新
    Module2.Date_Country
    
    'ピポットテーブルの更新
    
    'msgの表示
    Dim msg As String
    msg = "更新が完了しました。" & vbCrLf
    If after_datacount - datacount = 0 Then
        msg = msg & "新しいレコードはありません"
    Else
        msg = msg & after_datacount - datacount & " 件追加されました"
    End If
    
    Worksheets("Dashboard").Select
    
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
    MsgBox "Error!! " & Err.Description & "APIサーバーをは正常ですか？"
    request = 3
    Set HttpReq = Nothing 'オブジェクト解放

cleanUP:
    Set HttpReq = Nothing 'オブジェクト解放
    
End Function
