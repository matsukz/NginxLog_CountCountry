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

    CountryEndX = 2
    
    'データを消す
    Dim AGC_Y As Integer: AGC_Y = 2
    While Not AGC.Cells(AGC_Y, 1) = ""
        AGC.Cells(AGC_Y, 1) = "": AGC.Cells(AGC_Y, 2) = ""
        AGC_Y = AGC_Y + 1
    Wend
    
    While Not DC.Cells(1, CountryEndX) = ""

        Check = DC.Cells(1, CountryEndX)
        Debug.Print (Check)
    
        'Date_Countryのシートで集計する
        sum = 0: CountEndY = 2
        
        '各列のデータをsumに加算する
        While Not DC.Cells(CountEndY, CountryEndX) = ""
            sum = sum + DC.Cells(CountEndY, CountryEndX)
            CountEndY = CountEndY + 1
        Wend
        
        Debug.Print (sum)
        
        If sum <= 0 Then
            'sumが0以下なら結果を出力しない
        ElseIf sum <= 50 Then
            '5以下はその他にする
            sum_other = sum_other + sum
        Else
            'withステートメントが効かない謎
            AGC.Cells(AG_CountY, AG_CountX) = Check
            AGC.Cells(AG_CountY, AG_CountX + 1) = sum
             '出力先セル番地を移動させておく
            AG_CountX = 1: AG_CountY = AG_CountY + 1
        End If
        
        CountryEndX = CountryEndX + 1
    Wend
    
    '割合を降順に並び替える
    Module2.csort
    
    'ソート後にその他を追記する
    AGC.Cells(AG_CountY, AG_CountX) = "その他"
    AGC.Cells(AG_CountY, AG_CountX + 1) = sum_other
    
End Sub

Sub csort()
    Dim AGC As Worksheet: Set AGC = Worksheets("AG_Date_Country")
    AGC.Sort.SortFields.Clear
    AGC.Sort.SortFields.Add2 Key:=Range( _
        "C2:C100"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With AGC.Sort
        .SetRange Range("A2:C100")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
