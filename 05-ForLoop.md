# For文の基礎的な文法

プログラムを作成していると、「何度も同じことをやりたい」という状況が頻繁に登場します。  

今回の目的である万年カレンダーの作成の場合で言えば、A列からG列までの間、前日の値に+1日した数値を入力するということを6回くりかえしたいというような事もそれにあたります。
![](images/05-ForLoop/05-ForLoop20221503-155153.png)

もちろん、これまでの章でやったことだけでも実現は可能です。  

たとえば

```vb
    Cells(1, 1).Value = 1
    Cells(1, 2).Value = Cells(1, 1).Value + 1
    Cells(1, 3).Value = Cells(1, 2).Value + 1
    Cells(1, 4).Value = Cells(1, 3).Value + 1
    Cells(1, 5).Value = Cells(1, 4).Value + 1
    Cells(1, 6).Value = Cells(1, 5).Value + 1
    Cells(1, 7).Value = Cells(1, 6).Value + 1
    Cells(2, 1).Value = Cells(1, 7).Value + 1
    Cells(2, 2).Value = Cells(2, 1).Value + 1
    Cells(2, 3).Value = Cells(2, 2).Value + 1
    Cells(2, 4).Value = Cells(2, 3).Value + 1
            ・
            ・
            ・
```
というようなことを全31行記入すればいいのですが、面倒です。  
また、今回はたまたま月のカレンダーであるため、書いてもせいぜい31行ですが、これが年間カレンダーだったり、全社員に対する処理だったりすると数百から千行以上同じことを書く必要がでてきてしまい、現実的ではありません。  

そのため、VBAには`For ～　Next構文`という繰り返し処理のためのルールが用意されています。

イメージとしては以下のようなものです。  

![](images/05-ForLoop/05-ForLoop20221503-161738.png)

プログラムは上から順々に流れてきますが、For ～ Next構文(図の緑の場所)に入ると、For で定義された回数分、処理が繰り返し行われます。

