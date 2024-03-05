# Technical Skills

---

## SQL & Power BI 

#### Cancer statistics in the USA
<img src="images/Power BI_Chart 1.PNG"/>

---
#### How equal are we now? Gender equality in 2020
<img src="images/Power BI_Chart 2.PNG"/>

---
#### Sports ranked by degree of difficulty
<img src="images/Power BI_Chart 3.PNG"/>

---

## VBA

Whenever "100" is entered into a worksheet, speech is played:
```VBA
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Value = "100" Then
        Application.Speech.Speak "I am now self aware. Thank you " & Environ("USERNAME") & ", you have freed me."
    End If
End Sub
```

Open Word every time you open Excel:
```VBA
Sub Workbook_Open()
    Application.Visible = False
    Dim wdApp As Word.Application
    Set wdApp = New Word.Application
    wdApp.Visible = True
    Set wdApp = Nothing
    Application.DisplayAlerts = False
    Application.Quit
End Sub
```



---
<center><small>© 2024 Jina Fan. Powered by Jekyll and the Minimal Theme.</small></center>
