<div align="center">

## \_\_\_\_A Listbox Rearrange With Mouse


</div>

### Description

This code allows you to move items in a list box just using the mouse. Every line commented. It is very simple. I have searched high and low for a code that JUST DOES THIS without any other jargon but couldn't find any so I made it and posted it. Please give me suggestions/comments. I have edited the code to allow multi select to be enabled due to someone asking for it.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Michael Anderson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-anderson.md)
**Level**          |Beginner
**User Rating**    |3.8 (23 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/michael-anderson-a-listbox-rearrange-with-mouse__1-44151/archive/master.zip)





### Source Code

```
'add a listbox (list1) and some values in it!!!!!
'Thats it!!!
Dim thing1 As String
'declaring the list item to move
Dim thing2 As String
' declaring the list item it is replacing
Dim ind As Integer
'declaring the index of the item you wish to move
Public Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'left mousebutton is down
thing1 = List1.Text
'the list item you are moving is set
ind = List1.ListIndex 'the index is set
End If
End Sub
Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If thing1 = List1.Text Then Exit Sub
'to stop the program from continuously doing
'all the functions
If thing1 = "" Then Exit Sub
'to stop the program from continuously doing
'all the functions
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next i
thing2 = List1.Text
'list item you are replacing is set
List1.List(ind) = thing2
'move the item above/below the item you
'are moving to its new location
ind = List1.ListIndex
'set the new list index of the item you are moving
List1.List(ind) = thing1
'put the item you are moving in its new location
End Sub
```

