<div align="center">

## a Credit Card Number validation


</div>

### Description

Ever need to see if a credit card number is valid? Well, here is your chance. I did NOT write this code, I found it on the web!! Also, this may tell you if a number is valid, not if it works
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dustin Davis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dustin-davis.md)
**Level**          |Intermediate
**User Rating**    |2.3 (7 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dustin-davis-a-credit-card-number-validation__1-5696/archive/master.zip)





### Source Code

```
Function CheckCard(CCNumber As String) As Boolean
  Dim Counter As Integer, TmpInt As Integer
  Dim Answer As Integer
  Counter = 1
  TmpInt = 0
  While Counter <= Len(CCNumber)
    If (Len(CCNumber) Mod 2) Then
      TmpInt = Val(Mid$(CCNumber, Counter, 1))
      If Not (Counter Mod 2) Then
        TmpInt = TmpInt * 2
        If TmpInt > 9 Then TmpInt = TmpInt - 9
      End If
      Answer = Answer + TmpInt
      Counter = Counter + 1
    Else
      TmpInt = Val(Mid$(CCNumber, Counter, 1))
      If (Counter Mod 2) Then
        TmpInt = TmpInt * 2
        If TmpInt > 9 Then TmpInt = TmpInt - 9
      End If
      Answer = Answer + TmpInt
      Counter = Counter + 1
    End If
  Wend
  Answer = Answer Mod 10
  If Answer = 0 Then CheckCard = True
End Function
```

