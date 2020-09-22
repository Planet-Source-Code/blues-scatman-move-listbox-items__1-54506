<div align="center">

## Move listbox items


</div>

### Description

I was asked how to move items in a listbox up and down, to place then in any order of my choice.

Here's one way of doing it. You need two buttons, I named them 'buttonUp' and 'buttonDown', and ofcourse a listbox named 'List1'. Then all you have to do is select the item you wish to move, and press the up or down button. Great for playlists.

Scatman
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Blues Scatman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/blues-scatman.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/blues-scatman-move-listbox-items__1-54506/archive/master.zip)





### Source Code

```
<pre>
Private Sub buttonDown_Click()
 Dim nItem As Integer
 With list1
 If .ListIndex < 0 Then Exit Sub
 nItem = .ListIndex
 If nItem = .ListCount - 1 Then Exit Sub
 .AddItem .Text, nItem + 2
 .RemoveItem nItem
 .Selected(nItem + 1) = True
 End With
End Sub
'----------------------------------------
Private Sub ButtonUp_Click()
 Dim nItem As Integer
 With list1
 If .ListIndex < 0 Then Exit Sub
 nItem = .ListIndex
 If nItem = 0 Then Exit Sub
 .AddItem .Text, nItem - 1
 .RemoveItem nItem + 1
 .Selected(nItem - 1) = True
 End With
End Sub
</pre>
```

