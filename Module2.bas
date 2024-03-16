'https://github.com/matsukz/NginxLog_CountCountry

Sub Date_Country()
    Dim CountEndY As Integer
    Dim Check As String '国コードを格納する
    Dim sum As Integer: Dim sum_other As Integer '各国の合計とその他の集計
    sum = 0: sum_other = 0
    
    '使用するシートを変数に格納しておく
    Dim DC As Worksheet: Set DC = Worksheets("Date_Country")
    Dim AGC As Worksheet: Set AGC = Worksheets("AG_Date_Country")
    
    'グラフに含めるセルの番地(国コード)
    Dim AG_CountX, AG_CountY As Integer
    AG_CountX = 1: AG_CountY = 2

    '「その他」の基準(なければ0をセットする)
    If AGC.Range("I1") = "" Then AGC.Range("I1") = 0
    Dim other As Integer: other = AGC.Range("I1")
    
    'データを消す
    Dim AGC_Y As Integer: AGC_Y = 2
    Dim i As Byte 'ループ用
    While Not AGC.Cells(AGC_Y, 1) = ""
        i = 1
        For i = 1 To 3 Step 1
            AGC.Cells(AGC_Y, i) = ""
        Next i
        AGC_Y = AGC_Y + 1
    Wend
    
    Dim CountryEndX As Integer: CountryEndX = 2
    While Not DC.Cells(1, CountryEndX) = ""

        Check = DC.Cells(1, CountryEndX)
        
        If Check = "JP" And AGC.Range("I2") = False Then
            '日本を含めない場合は以下の加算表示処理をスキップする
        Else
    
            'Date_Countryのシートで集計する
            sum = 0: CountEndY = 2
        
            '各列のデータをsumに加算する
            While Not DC.Cells(CountEndY, CountryEndX) = ""
                sum = sum + DC.Cells(CountEndY, CountryEndX)
                CountEndY = CountEndY + 1
            Wend
        
            If sum <= 0 Then
                'sumが0以下なら結果を出力しない
            ElseIf sum <= other Then
                '5以下はその他にする
                sum_other = sum_other + sum
            Else
            
            AGC.Cells(AG_CountY, AG_CountX) = Check
            AGC.Cells(AG_CountY, AG_CountX + 1) = sum
            '出力先セル番地を移動させておく
            AG_CountX = 1: AG_CountY = AG_CountY + 1
            
            End If
        
        End If
        
        CountryEndX = CountryEndX + 1
        
    Wend
    
    '割合を求めて、その他を後で使う
    Dim ratio As Single
    ratio = Module2.ratio(sum_other)
    
    '割合を降順に並び替える(並び替える範囲を引数にする)
    Module2.csort Range(AGC.Cells(2, 1), AGC.Cells(AG_CountY - 1, 3)).address, _
                  Range(AGC.Cells(2, 3), AGC.Cells(AG_CountY - 1, 3)).address
    
    'ソート後にその他を追記する。その他が0のときは無視
    If sum_other = 0 Then
        '何もしない
    Else
        With AGC
            .Cells(AG_CountY, AG_CountX) = "その他"
            .Cells(AG_CountY, AG_CountX + 1) = sum_other
            .Cells(AG_CountY, AG_CountX + 2) = ratio
        End With
    End If
    
End Sub

Function ratio(ByVal sum_other As Integer) As Single
    Dim AGC As Worksheet: Set AGC = Worksheets("AG_Date_Country")
    Dim sum_all As Integer: sum_all = 0
    Dim AGC_Y As Integer: AGC_Y = 2
    
    '合計を求める
    While Not AGC.Cells(AGC_Y, 1) = ""
        sum_all = sum_all + AGC.Cells(AGC_Y, 2)
        AGC_Y = AGC_Y + 1
    Wend
    
    sum_all = sum_all + sum_other

    '割合を求める
    AGC_Y = 2
    While Not AGC.Cells(AGC_Y, 1) = ""
        AGC.Cells(AGC_Y, 3) = AGC.Cells(AGC_Y, 2) / sum_all
        AGC_Y = AGC_Y + 1
    Wend
    
    'その他の割合を返す
    ratio = sum_other / (sum_all)
    
End Function

Sub csort(ByVal sort_range As String, ByVal sort_key As String)
    Dim AGC As Worksheet: Set AGC = Worksheets("AG_Date_Country")
    AGC.Sort.SortFields.Clear
    AGC.Sort.SortFields.Add2 Key:=Range( _
        sort_key), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With AGC.Sort
        .SetRange Range(sort_range)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
