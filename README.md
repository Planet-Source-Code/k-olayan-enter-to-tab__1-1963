<div align="center">

## Enter to Tab


</div>

### Description

This code gives the "enter" key the same functionality as the "tab" key in a vb form. When the user presses the "enter" key, it moves the focus to the next control, based on the tab index order. Don't forget to set the KeyPreview property of the form to True. (My thanks to Pennington. His TabToEnter2 code laid the foundation for this code.)
 
### More Info
 
The KeyPreview event of the form must be set to true for the Form_KeyPress event to fire.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[K\.Olayan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/k-olayan.md)
**Level**          |Unknown
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/k-olayan-enter-to-tab__1-1963/archive/master.zip)





### Source Code

```
'add this to your form's code
Private Sub Form_KeyPress(KeyAscii As Integer)
 'catch both "Enter" keys on keyboard
 If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeySeparator) Then
  SendKeys "{tab}"
 End If
End Sub
```

