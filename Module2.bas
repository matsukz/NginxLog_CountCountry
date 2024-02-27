Attribute VB_Name = "Module2"
Sub Date_Country()
    Dim CountryEndY, CountEndX, CountEndY As Integer
    Dim Check As String
    Dim sum As Integer
    
    CountryEndY = 2
    
    CountEndX = 2: CountEndY = 2
    
    Worksheets("AG_Date_Country").Select
    
    While Not Cells(CountryEndY, 1) = ""
        Worksheets("AG_Date_Country").Select: Check = Cells(CountryEndY, 1)
    
        'Date_Countryのシートで集計する
        Worksheets("Date_Country").Select
        sum = 0: CountEndY = 2
        
        '各列のデータをsumに加算する
        While Not Cells(CountEndY, CountEndX) = ""
            sum = sum + Cells(CountEndY, CountEndX)
            CountEndY = CountEndY + 1
        Wend
        
        'sumを反映する
        Worksheets("AG_Date_Country").Select
        Cells(CountryEndY, 2) = sum
        
        CountEndX = CountEndX + 1
        CountryEndY = CountryEndY + 1
    Wend
        
    Worksheets("Dashboard").Select
    
End Sub

