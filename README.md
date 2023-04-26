# VBA_data_to_clean_control
研究の効率化　1

## Nichino
```xlsm
Sub 表示形式変更()
 
' 表示形式変更のコード
    Dim dt As String
    dt = "m/d "" ""h:mm"
    
    ' 変更したい列を指定して、その列の範囲を取得する
    Dim targetRange As Range
    Set targetRange = Range("C:C")
    
    ' 取得した範囲に対して、表示形式を変更する
    targetRange.NumberFormat = dt


'その左隣のセルに分を抜き出すコード
    




'そのさらに左隣に時間を表示するコード




End Sub
```

### Nishino

#### NIshino