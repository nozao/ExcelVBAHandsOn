# Excel VBA ハンズオン講義

万年カレンダーの作成を通してExcelマクロ(VBA)の最初の一歩を学ぶための資料です。

![](images/README/README20211712-103622.png)

最終的にはこれが完成します

## 対象者

今まで一度もプログラムに触れたことがない人  
私有PCでアプリ版のExcel(2016以降)が使用できる人

## 学習範囲

VBAの基本的な文法と、マクロ開発の流れを学習することができます。  
VBAもプログラム言語の一つであるため、ここで学習した文法や考え方は他の言語でも応用することが可能です。

1. Excel(アプリ版)でVBA開発をするための環境構築
2. まずは動かしてみる-Hello,World作成(指定したセルに任意の文字を表示する)
3. Cellsについて
4. セルに入力された文字を取得する
5. 入力された文字を受け取って計算し、他のセルに書き出す(変数について)
6. For文の基礎的な文法
7. If文の基礎的な文法
8. 用意されている関数を使う
9. オリジナルの関数を作る
10. ボタンと連動させる

本編とは別に、下記の内容についても説明します

- デバッグの仕方(ブレイクポイント、ステップ実行、ウォッチウィンドウ)
- エラーの対処法
- プログラム作成のコツ(ロジックの作り方、ちょっとづつ作ってみる)
- イベントドリブンについて


## ソースコード

参考までに、このマクロのプログラムコードはこうなっています。  
これを詳細に解説していきます。

```vb
Sub main()
    Call DrawCalendar(Cells(1, 3).Value, Cells(1, 4).Value)
End Sub

Function DrawCalendar(TargetYear As Integer, TargetMonth As Integer)

    Dim FirstDate As Date
    Dim PrintDate As Date
    Dim DiffDay As Integer
    
    FirstDate = DateSerial(TargetYear, TargetMonth, 1)
    DiffDay = Weekday(FirstDate, vbSunday)
    
    PrintDate = FirstDate - DiffDay + 1
    
    For w = 3 To 8
        For d = 1 To 7
            If Month(PrintDate) = Month(FirstDate) Then
                Cells(w, d).Value = PrintDate
            Else
                Cells(w, d).Value = ""
            End If
            PrintDate = PrintDate + 1
        Next
    Next

End Function
```