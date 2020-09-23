<div align="center">

## \*Improved\* Cool Form Close


</div>

### Description

'This is a much improved version of the cool form close code submitted by Jas Batra. It shrinks the First of all, it is in function form, second it is a 'lot' faster and smoother. Code is fully documented for beginners.
 
### More Info
 
'Inputs are:

' coolCloseForm closeForm,speed

'closeform is the form to close

'speed is anything from 1 to about 100 or more...

'Completely documented for beginners.

'Nothing

'None so far, make sure that no data will be lost, because when this code runs, it unloads the form.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jono Spiro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jono-spiro.md)
**Level**          |Unknown
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jono-spiro-improved-cool-form-close__1-1874/archive/master.zip)

### API Declarations

```
'None
```


### Source Code

```
'If you want to try this code in action:
' make a new project and add a module
'double click on the form and add the following code:
Private Sub Form_Load()
 Form1.Height = 6400
 Form1.Width = 10000
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbRightButton Then
 coolCloseForm Me, 20
 Else
 Dim a As New Form1
 a.Height = a.Height / 2
 a.Width = a.Width / 2
 a.Show
 End If
End Sub
'Then add the coolCloseForm code to the module
'Now run the program, left click a few times to add new forms to screen, and then right click on them to make them go away.
'END OF EXAMPLE CODE
'
'
'
'ALL CODE BELOW TO THE BOTTOM IS THE ACTUAL MODULE CODE, ABOVE CODE IS ALL OPTIONAL!!
'
Public Function coolCloseForm(closeForm As Form, speed As Integer)
 'make sure speed is more than 1
 If speed = 0 Then
 MsgBox "Speed cannot zero"
 Exit Function
 End If
 'closeform is the form to close
 'speed is anything from 1 to about 100
 On Error Resume Next
 'set the scalemode to twips so that the do statements will work
 closeForm.ScaleMode = 1
 'so the code wont crash
 closeForm.WindowState = 0
 'do until the height is the minimum possible
 Do Until closeForm.Height <= 405
 'let the computer process
 DoEvents
 'make the form shorter by the speed * 10
 closeForm.Height = closeForm.Height - speed * 10
 'make the top of the form lower by the speed * 5
 closeForm.Top = closeForm.Top + speed * 5
 Loop
 'do until the width is the minimum possible
 Do Until closeForm.Width <= 1680
 'let the computer process
 DoEvents
 'make the form less wide by the speed * 10
 closeForm.Width = closeForm.Width - speed * 10
 'make the left of the form farther to the righ by the speed * 5
 closeForm.Left = closeForm.Left + speed * 5
 Loop
 'when its all done, unload the form
 Unload closeForm
End Function
```

